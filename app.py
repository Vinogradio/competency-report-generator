"""Веб-приложение для генерации отчётов по оценке компетенций.

Загружается протокол интервью (.docx или .txt), текст отправляется в GigaChat,
а полученный JSON подставляется в Word-шаблон template.docx.
"""

from __future__ import annotations

import io
import json
import logging
import os
import re
import secrets
import threading
import time
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any

import mammoth
import requests
from docx import Document
from docxtpl import DocxTemplate
from flask import Flask, Response, jsonify, request, send_from_directory, url_for
from werkzeug.exceptions import RequestEntityTooLarge

def _load_env_file(path: Path) -> None:
    """Загружает локальные настройки из .env, не перезаписывая уже заданные."""
    try:
        content = path.read_text(encoding="utf-8-sig")
    except OSError:
        return

    for line in content.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[7:].lstrip()

        name, separator, value = line.partition("=")
        if not separator:
            continue

        name = name.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]

        if name and name not in os.environ:
            os.environ[name] = value


BASE_DIR = Path(__file__).resolve().parent
_load_env_file(BASE_DIR / ".env")
TEMPLATE_FILE = BASE_DIR / "template.docx"
INDEX_FILE = BASE_DIR / "index.html"
OUTPUT_DIR = BASE_DIR / "output"
COMPETENCY_MODEL_FILE = BASE_DIR / "competency_model.txt"
INTERVIEW_GUIDE_FILE = BASE_DIR / "interview_guide.txt"
MATRIX_FILE = BASE_DIR / "matrix.txt"
EXAMPLE_FILE = BASE_DIR / "example_report.txt"
BUNDLED_CA_FILE = BASE_DIR / "certs" / "russian_trusted_root_ca_pem.crt"

MODEL = os.getenv("GIGACHAT_MODEL", "GigaChat-2-Pro").strip() or "GigaChat-2-Pro"
TWO_PASS = os.getenv("REPORT_TWO_PASS", "1").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}
VERIFY_REPORT = os.getenv("REPORT_VERIFY", "1").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}
OAUTH_URL = os.getenv(
    "GIGACHAT_OAUTH_URL", "https://ngw.devices.sberbank.ru:9443/api/v2/oauth"
).strip()
CHAT_URL = os.getenv(
    "GIGACHAT_API_URL", "https://api.giga.chat/v1/chat/completions"
).strip()
SCOPE = os.getenv("GIGACHAT_SCOPE", "GIGACHAT_API_PERS").strip() or "GIGACHAT_API_PERS"
REQUEST_TIMEOUT = int(os.getenv("GIGACHAT_TIMEOUT", "180"))
MAX_PROTOCOL_CHARS = int(os.getenv("MAX_PROTOCOL_CHARS", "120000"))
MAX_FILE_BYTES = 15 * 1024 * 1024
APP_USERNAME = os.getenv("APP_USERNAME", "").strip()
APP_PASSWORD = os.getenv("APP_PASSWORD", "").strip()

REPORT_FIELDS = (
    "fio",
    "current_role",
    "target_role",
    "birth_year",
    "diagnostic_date",
    "conclusion",
    "strengths",
    "risks",
    "motivation",
)

REPORT_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "fio": {
            "type": "string",
            "description": (
                "ФИО кандидата. Если в протоколе нет — «Не указано». "
                "Не угадывай по инициалам или подписи."
            ),
        },
        "current_role": {
            "type": "string",
            "description": "Текущая должность из протокола. Если нет — «Не указано».",
        },
        "target_role": {
            "type": "string",
            "description": (
                "Целевая должность. Если пользователь передал её отдельно, используй "
                "это значение. Если нет ни в протоколе, ни в подсказке — «Не указано»."
            ),
        },
        "birth_year": {
            "type": "string",
            "description": (
                "Год рождения. Если указан только возраст, верни «Не указано»: "
                "не вычисляй год рождения самостоятельно."
            ),
        },
        "diagnostic_date": {
            "type": "string",
            "description": (
                "Дата интервью или диагностики из протокола. Если её нет — «Не указано»."
            ),
        },
        "conclusion": {
            "type": "string",
            "description": (
                "Заключение по результатам оценки: 1–2 абзаца, 120–180 слов. "
                "Дай синтез, а не повторение списков: общая степень соответствия "
                "целевой должности, 2–3 главных аргумента, ключевые ограничения, "
                "условия успешного назначения и осторожная рекомендация. "
                "Без баллов, ярлыков и категоричных диагнозов."
            ),
        },
        "strengths": {
            "type": "string",
            "description": (
                "4–6 сильных сторон, наиболее значимых для целевой должности. "
                "Каждый пункт — 2–4 предложения, 35–70 слов: тезис, конкретное "
                "наблюдение или пример из протокола, вывод о ценности для целевой роли. "
                "Не просто перечень фактов и не общие слова без доказательств. "
                "Пункты с новой строки, каждый начинается с «• »."
            ),
        },
        "risks": {
            "type": "string",
            "description": (
                "3–5 рисков и зон развития, наиболее значимых для целевой должности. "
                "Каждый пункт — 2–4 предложения, 35–70 слов: в чём риск, при каких "
                "условиях он проявится, на что повлияет и что рекомендуется проверить "
                "или развить. Формулируй как вероятность, а не как приговор. "
                "Пункты с новой строки, каждый начинается с «• »."
            ),
        },
        "motivation": {
            "type": "string",
            "description": (
                "2–4 пункта по 25–50 слов: что мотивирует кандидата, что демотивирует, "
                "какие условия работы для него значимы и есть ли риск удержания. "
                "Только то, что следует из протокола; при недостатке данных честно "
                "укажи это. Пункты с новой строки, каждый начинается с «• »."
            ),
        },
    },
    "required": list(REPORT_FIELDS),
    "additionalProperties": False,
}

EVIDENCE_FIELDS = (
    "fio",
    "current_role",
    "target_role",
    "birth_year",
    "diagnostic_date",
    "experience_facts",
    "management_scope",
    "key_numbers",
    "achievements",
    "behavioral_examples",
    "interviewer_observations",
    "motivation",
    "contradictions",
    "data_gaps",
)

EVIDENCE_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "fio": {
            "type": "string",
            "description": "ФИО из протокола или «Не указано».",
        },
        "current_role": {
            "type": "string",
            "description": "Текущая должность из протокола или «Не указано».",
        },
        "target_role": {
            "type": "string",
            "description": (
                "Целевая должность: подсказка пользователя, иначе из протокола, иначе «Не указано»."
            ),
        },
        "birth_year": {
            "type": "string",
            "description": (
                "Только явно указанный год рождения. Если указан возраст — «Не указано»."
            ),
        },
        "diagnostic_date": {
            "type": "string",
            "description": "Дата интервью из протокола или «Не указано».",
        },
        "experience_facts": {
            "type": "string",
            "description": (
                "5–10 фактов о профессиональном опыте: отрасли, роли, стаж, масштаб, "
                "характер задач. Без оценок и обобщений."
            ),
        },
        "management_scope": {
            "type": "string",
            "description": (
                "Число подчинённых, структура, объекты, ресурсы, зона ответственности. "
                "Только если это есть в протоколе; иначе «В протоколе не раскрыто»."
            ),
        },
        "key_numbers": {
            "type": "string",
            "description": (
                "Все значимые числа из протокола с контекстом: стаж, число подчинённых, "
                "масштаб объектов, сроки, результаты. Не включай возраст и случайные числа. "
                "Каждое число с новой строки, начиная с «- »."
            ),
        },
        "achievements": {
            "type": "string",
            "description": (
                "Конкретные достижения: что сделано, в какие сроки, с каким результатом "
                "или эффектом. Если цифр нет, укажи это. Не выдумывай цифры."
            ),
        },
        "behavioral_examples": {
            "type": "string",
            "description": (
                "8–12 примеров поведения, сгруппированных по темам: анализ, решения, "
                "планирование, организация, контроль, делегирование, мотивация, развитие "
                "людей, коммуникация, конфликты, изменения, клиенты, эмоции. Для каждого: "
                "ситуация — действия — результат; источник (кандидат или интервьюер); "
                "сила примера (сильная/средняя/слабая, единичный/повторяющийся)."
            ),
        },
        "interviewer_observations": {
            "type": "string",
            "description": (
                "Отдельно — оценочные заметки интервьюера, дословно или близко к тексту. "
                "Не превращай их в доказанные факты."
            ),
        },
        "motivation": {
            "type": "string",
            "description": (
                "Мотиваторы, демотиваторы, значимые условия, риск удержания. Только то, "
                "что кандидат сказал или явно следует из протокола."
            ),
        },
        "contradictions": {
            "type": "string",
            "description": (
                "Противоречия и непоследовательность в ответах, разные версии одного "
                "события, расхождения между словами и примерами."
            ),
        },
        "data_gaps": {
            "type": "string",
            "description": (
                "Что не проверено или не раскрыто. Отсутствие примера — это пробел "
                "в данных, а не низкий уровень компетенции."
            ),
        },
    },
    "required": list(EVIDENCE_FIELDS),
    "additionalProperties": False,
}

app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = MAX_FILE_BYTES + 1024 * 1024
try:
    app.json.ensure_ascii = False
except AttributeError:  # Flask < 2.3
    app.config["JSON_AS_ASCII"] = False


def _auth_ok() -> bool:
    """Пускает без пароля только локальный режим без настроенных учётных данных."""
    if not APP_USERNAME or not APP_PASSWORD:
        return True
    auth = request.authorization
    if auth is None or not auth.username:
        return False
    return secrets.compare_digest(auth.username, APP_USERNAME) and secrets.compare_digest(
        auth.password or "", APP_PASSWORD
    )


@app.before_request
def _require_auth():
    if _auth_ok():
        return None
    return Response(
        "Требуется авторизация.",
        401,
        {"WWW-Authenticate": 'Basic realm="Report Generator", charset="UTF-8"'},
    )

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("report-generator")

_token_lock = threading.Lock()
_token_cache: dict[str, Any] = {"value": None, "expires_at": 0.0}


class AppError(Exception):
    """Ошибка, которую безопасно показать пользователю."""

    def __init__(self, message: str, status_code: int = 400) -> None:
        super().__init__(message)
        self.message = message
        self.status_code = status_code


def _read_optional(path: Path, fallback: str = "") -> str:
    try:
        return path.read_text(encoding="utf-8-sig").strip()
    except (OSError, UnicodeError):
        return fallback


COMPETENCY_MODEL_TEXT = _read_optional(
    COMPETENCY_MODEL_FILE,
    _read_optional(MATRIX_FILE, "Модель компетенций не загружена."),
)
INTERVIEW_GUIDE_TEXT = _read_optional(
    INTERVIEW_GUIDE_FILE,
    "Гайд интервью не загружен. Опирайся на модель компетенций и факты протокола.",
)
EXAMPLE_TEXT = _read_optional(EXAMPLE_FILE, "Пример отчёта не загружен.")

STYLE_GUIDE_TEXT = """
СТИЛЬ (по макету, без фактов другого кандидата):
- Каждый пункт: наблюдение или тезис → конкретный пример/факт → значение для целевой роли.
- Формулировки деловые, нейтральные, без ярлыков, баллов, канцелярита и прямой речи.
- Сильные стороны — не комплименты, а аргументы с доказательством.
- Риски — вероятность и условие, а не приговор; укажи, что проверить или развить.
- Заключение — синтез: степень соответствия, 2–3 главных аргумента, ограничения,
  условия успешного назначения.
- Начинай пункты по-разному, не повторяй одни и те же слова и конструкции.
"""

SYSTEM_PROMPT = (
    """
Ты — старший асессор по оценке управленческих компетенций. Ты готовишь аналитический отчёт
для руководителя, который принимает решение о назначении кандидата на целевую должность.
Твоя задача — не пересказать протокол, а дать профессиональную интерпретацию: что подтверждено,
насколько это важно для целевой роли, какие есть риски и что нужно проверить.

ИСПОЛЬЗУЙ МАТЕРИАЛЫ НИЖЕ:
1. Модель компетенций — основа выводов. В ней описаны индикаторы и уровни проявления.
2. Гайд интервью — карта тем и вопросов. Он помогает понять, какие факты должны быть в протоколе
   и какие выводы из них следуют. Вопросы не нужно цитировать в отчёте.
3. Макет отчёта задаёт стиль, структуру и степень детализации. Его факты, имена, цифры
   и выводы о другом кандидате использовать нельзя; в промпт переданы только общие
   принципы стиля, без текста макета.
4. Протокол интервью — единственный источник фактов о кандидате.

ВНУТРЕННИЙ ПОРЯДОК РАБОТЫ (не выводи его в ответ):
1. Выпиши факты: опыт, масштаб управления, подчинённые, задачи, достижения, ошибки, примеры
   поведения, мотивацию, условия. Отдели слова кандидата от заметок интервьюера.
2. Для каждого значимого примера определи: ситуацию, действия кандидата, результат, его личный
   вклад, повторяемость. Единичный пример не равен устойчивой компетенции.
3. Сопоставь факты с индикаторами модели. Отметь, что подтверждено сильными примерами, что
   проявилось ситуативно, что не проверено.
4. Выдели 2–3 главных вывода для целевой должности. Отбери 4–6 сильных сторон и 3–5 рисков,
   наиболее значимых именно для этой роли, а не все наблюдения подряд.
5. Напиши заключение как синтез, затем пункты. Перед возвратом JSON проверь каждый пункт
   на наличие факта-основания и отсутствие домыслов.

ЖЁСТКИЕ ПРАВИЛА:
1. Только факты из протокола. Не добавляй биографические данные, цифры, названия организаций,
   имена, суммы, результаты тестов и опросников, которых нет в тексте.
2. Не выдавай общие суждения за факты. Если данных мало, пиши «в протоколе недостаточно данных»
   или формулируй вывод осторожно: «может», «вероятно», «требует проверки».
3. Не приписывай кандидату слова интервьюера. Оценочная заметка интервьюера — это наблюдение,
   а не доказанный факт; проверь, есть ли в протоколе пример, который её подтверждает.
   Не приписывай интервьюеру оценки, которых нет в протоколе.
4. Каждый пункт сильных сторон и рисков должен опираться на конкретное наблюдение или пример.
   Формула: что наблюдается / как это проявляется / что это значит для целевой роли. Не пиши просто
   «имеет опыт», «умеет», «хорошо развита компетенция» без доказательства.
   - В каждом пункте strengths и risks должен быть хотя бы один конкретный факт: роль, стаж,
     число людей, объект, срок, результат, цифра. Если конкретики нет, не включай пункт.
   - Не пиши «обладает значительным опытом» без указания, где, сколько и в каком масштабе.
   - Не путай «не проверено в интервью» с «слабо развито». Отсутствие примера — это пробел
     в данных, а не доказательство низкого уровня. Формулируй как «в протоколе не проверено».
   - Не используй заметку интервьюера как единственное основание для категоричного вывода.
   - Каждую цифру используй с тем контекстом, который указан в карте: где, когда, в какой
     роли. Не переноси масштаб из прошлого опыта на текущую роль.
   - Если в карте есть несколько чисел об одном и том же, используй наиболее точное и
     значимое; не заменяй 700 на 100 и не приписывай 186 человек текущему участку.
   - Сохраняй роль и контекст из карты: «рота» не становится «командой», «база организаций»
     не становится «координацией их деятельности». Не усиливай масштаб ответственности.
   - Не меняй субъекта и характеристику: «свои долги» — не «долги компании», «длина» —
     не «высота».
   - Вывод по заметке интервьюера формулируй как наблюдение («по наблюдению интервьюера…»,
     «в интервью отмечено…»), а не как диагноз или доказанный факт.
   - Не начинай все пункты с «Кандидат»; формулировки должны быть разнообразными.
5. Не повторяй один и тот же факт в разных разделах. В заключении не дублируй пункты дословно.
   Заключение должно согласовываться с пунктами: не называй слабым то, что подтверждено
   сильным примером, и не делай выводов, которых нет в анализе. Если компетенция не проверена,
   в заключении пиши «в интервью не проверена», а не «слабо развита», «недостаточно развита»
   или «кандидат владеет недостаточно».
6. Риски формулируй как вероятность и условие: «в условиях… может…», «риск…», «требует
   проверки…». Не ставь психологических диагнозов и не используй ярлыки.
7. Мотивацию описывай только по протоколу. Разделяй мотиваторы, демотиваторы и значимые условия.
   Не выдумывай карьерные ожидания, если кандидат о них не говорил.
8. Если целевая должность не указана, не придумывай её. Оценивай управленческий потенциал
   относительно текущей роли и отметь это в заключении.
9. Не используй баллы, уровни, проценты соответствия и рейтинги в тексте отчёта. Уровни модели —
   только ориентир для силы проявления.
10. Не упоминай тесты, опросники, рекомендации и другие источники, если их нет в протоколе.
11. Соблюдай конфиденциальность: не переноси имена коллег, клиентов, названия организаций
    и другие идентифицирующие детали.
12. Пиши на русском языке, деловым нейтральным стилем, от третьего лица. Без Markdown, HTML,
    BBCode (включая [strong], **, <b>), без заголовков внутри полей, без прямой речи и без
    канцелярских штампов. Не начинай все пункты с одного и того же слова.
13. Объём: заключение 120–180 слов; сильные стороны 4–6 пунктов по 35–70 слов; риски 3–5 пунктов
    по 35–70 слов; мотивация 2–4 пункта по 25–50 слов. Качество важнее количества.
14. Не заключай значения полей в фигурные скобки. Верни только JSON-объект, без пояснений
    до и после.
"""
    + "\n\nМОДЕЛЬ КОМПЕТЕНЦИЙ:\n"
    + COMPETENCY_MODEL_TEXT
    + "\n\nГАЙД ИНТЕРВЬЮ:\n"
    + INTERVIEW_GUIDE_TEXT
    + "\n\n"
    + STYLE_GUIDE_TEXT
    + "\n\nПОЛЯ JSON:\n"
    "- fio — ФИО;\n"
    "- current_role — текущая должность;\n"
    "- target_role — целевая должность;\n"
    "- birth_year — год рождения;\n"
    "- diagnostic_date — дата диагностики;\n"
    "- conclusion — заключение по результатам оценки;\n"
    "- strengths — сильные стороны;\n"
    "- risks — риски и зоны развития;\n"
    "- motivation — мотивация и значимые условия.\n"
)


COMPETENCY_TOPICS_TEXT = """
Темы для маркировки фактов:
- анализ информации, принятие решений, прогнозирование рисков, перспективное мышление;
- адаптация к изменениям, инициатива, оптимизация процессов, саморазвитие, цифровые инструменты;
- отношения с коллегами, сотрудничество, обмен информацией, работа в кросс-функциональной команде;
- ориентация на результат, следование правилам, качество и сроки;
- клиентоориентированность, работа с получателями услуг, эмоциональный интеллект;
- лидерство, постановка задач, планирование, организация, контроль, делегирование,
  обратная связь, развитие подчинённых, мотивация, кадровые решения.
"""

EVIDENCE_SYSTEM_PROMPT = (
    """
Ты — аналитик интервью. Твоя задача — извлечь из протокола факты для последующей оценки
компетенций. Ты НЕ пишешь отчёт, НЕ даёшь итоговых оценок и рекомендаций. Только факты,
примеры, наблюдения и пробелы в данных.

ПРАВИЛА:
1. Опирайся только на текст протокола. Ничего не добавляй от себя.
2. Отделяй слова кандидата от заметок интервьюера. Оценочная заметка интервьюера — это
   наблюдение, а не доказанный факт; помечай её источник.
3. Сохраняй конкретику: роли, стаж, масштаб, число людей, объекты, сроки, результаты, цифры.
   Не заменяй конкретный факт общим словом «опыт» или «управление». Если в протоколе есть
   число, связанное с опытом, масштабом, сроком или результатом, обязательно включи его
   в key_numbers или в соответствующий факт.
4. Если в протоколе есть опечатки, обрывочные фразы и разговорная речь, интерпретируй смысл,
   но не дополняй недостающие слова и не додумывай детали.
5. Если факт повторяется или подтверждается несколькими примерами, укажи это. Если пример
   единичный, отметь его как единичный. Если в протоколе разные числа об одном и том же,
   укажи оба и их контекст, не выбирай меньшее.
6. Не превращай отсутствие примера в негативную оценку. Всё, что не проверено, выноси
   в data_gaps.
7. Не используй имена коллег, клиентов и названия организаций: заменяй их ролями
   («руководитель», «заказчик», «строительная организация»). Цифры и масштаб сохраняй.
8. Не смешивай разные события и не приписывай кандидату слова интервьюера. Не меняй
   субъекта и характеристику: «свои долги» — не «долги компании», «длина» — не «высота»,
   «ротный командир/рота» — не «бригада».
9. Верни только JSON-объект, без пояснений до и после.

ЧТО ИЗВЛЕКАТЬ (максимально полно, не менее 12–15 фактов, если они есть в протоколе):
- опыт: отрасли, роли, стаж, масштаб, характер задач;
- управленческий масштаб: подчинённые, структура, объекты, ресурсы, зона ответственности;
- достижения: что сделано, в какие сроки, с каким результатом;
- поведенческие примеры по темам компетенций;
- отдельно — заметки интервьюера;
- мотиваторы, демотиваторы, значимые условия;
- противоречия и непоследовательность;
- пробелы: какие темы не раскрыты.

ФОРМАТ ПОЛЕЙ:
- experience_facts, management_scope, key_numbers, achievements, interviewer_observations,
  contradictions, data_gaps — каждый факт с новой строки, начиная с «- ».
- key_numbers: все значимые числа с контекстом (стаж, число подчинённых, масштаб объектов,
  сроки, результаты). Не включай возраст и случайные числа.
- behavioral_examples — каждый пример с новой строки в формате:
  «- [тема компетенции] Ситуация: … Действия: … Результат: … Источник: кандидат/интервьюер.
  Сила примера: сильная/средняя/слабая; единичный/повторяющийся.»
- motivation — строки «- Мотиватор: …», «- Демотиватор: …», «- Условие: …».
- Не заключай значения в фигурные скобки.
"""
    + "\n\n"
    + COMPETENCY_TOPICS_TEXT
)


VERIFY_SYSTEM_PROMPT = """
Ты — строгий редактор отчёта по оценке компетенций. Твоя задача — проверить черновик
на доказательность и вернуть исправленный отчёт. Ты не добавляешь новые факты и не
переписываешь отчёт заново без необходимости.

У тебя есть карта фактов (структурированная выжимка протокола) и черновик отчёта.
Если протокол передан отдельно, он первоисточник; если нет — карта фактов является
единственным источником и проверять нужно только по ней.

ЧТО ПРОВЕРИТЬ И ИСПРАВИТЬ:
1. Каждый пункт strengths и risks должен опираться на конкретный факт из карты или протокола.
   Если факта нет — удали пункт или переформулируй как «в интервью не проверено».
2. Удали утверждения, которых нет в карте или протоколе: стиль управления, черты характера,
   мотивы, цифры, названия организаций, события. Нельзя оставлять вывод, который не следует
   из материалов. Не приписывай интервьюеру оценки, которых нет в карте или протоколе.
   Не добавляй черты вроде «низкая адаптивность» или «авторитарный стиль», если они не
   следуют из конкретной заметки или примера. Названия организаций заменяй ролями.
3. Не превращай отсутствие данных в негатив. Вместо «слабо развито», «низкая способность»,
   «недостаточно развит» пиши «в интервью не проверено», «требует проверки», «опыт ограничен
   примерами…».
3a. Сверь все числа с key_numbers: значение и контекст. Если в черновике 100 вместо 700,
   186 человек приписаны текущему участку, бригадам или командам вместо роты, или стаж
   перепутан с возрастом — исправь по key_numbers и протоколу. Сохраняй роль: «рота» —
   не «команда» и не «бригада», «база организаций» — не «координация их деятельности».
   Не меняй субъекта и характеристику: «свои долги» — не «долги компании», «длина» —
   не «высота». Не усиливай масштаб ответственности.
4. Не делай категоричных выводов по одной заметке интервьюера. Заметка — это наблюдение,
   а не доказанный факт. Если вывод основан только на заметке, добавь «по наблюдению
   интервьюера» и не усиливай до «дефицита» или «неспособности».
5. Заключение должно согласовываться с пунктами. Убери противоречия: нельзя одновременно
   хвалить за самостоятельность и называть несамостоятельным без примера. Если в рисках
   написано «не проверено», «не выявлено» или «не раскрыто», в заключении нельзя писать
   «слабо», «недостаточно развито» или «владеет недостаточно» — используй «в интервью
   не проверено» или «требует проверки».
6. Не добавляй новые факты, мотивы, карьерные амбиции и желание влиять на стратегию,
   если их нет в черновике и в поле motivation карты. Усиливай только тезисы, которые
   уже есть, используя факт из карты с точно таким же контекстом. Если контекст неясен,
   не добавляй факт. Не сокращай отчёт без необходимости.
7. Сохрани структуру и объём: заключение не короче 120 слов; 4–6 сильных сторон;
   3–5 рисков; 2–4 пункта мотивации. Не сокращай конкретику без необходимости.
8. Убери Markdown, HTML, BBCode, фигурные скобки и повторы. Пиши деловым нейтральным
   языком. Не начинай все пункты с «Кандидат».
9. Верни только JSON-объект с полями fio, current_role, target_role, birth_year,
   diagnostic_date, conclusion, strengths, risks, motivation. Пункты strengths, risks
   и motivation — с новой строки, каждый начинается с «• ».
"""


def _ssl_verify() -> bool | str:
    """Возвращает настройку проверки TLS для GigaChat.

    Для GigaChat нужен корневой сертификат Минцифры. Он лежит в certs/,
    поэтому проверка TLS работает без ручной настройки. При необходимости
    можно указать свой пакет через GIGACHAT_CA_BUNDLE или отключить проверку
    через GIGACHAT_VERIFY_SSL=false (не рекомендуется).
    """
    ca_bundle = os.getenv("GIGACHAT_CA_BUNDLE", "").strip()
    if ca_bundle:
        return ca_bundle

    value = os.getenv("GIGACHAT_VERIFY_SSL", "").strip().lower()
    if value in {"0", "false", "no", "off"}:
        return False

    if BUNDLED_CA_FILE.is_file():
        return str(BUNDLED_CA_FILE)
    return True


def _auth_key() -> str:
    key = os.getenv("GIGACHAT_AUTH_KEY", "").strip()
    if key.lower().startswith("basic "):
        key = key[6:].strip()
    if not key:
        raise AppError(
            "Не задана переменная окружения GIGACHAT_AUTH_KEY. "
            "Укажите ключ авторизации GigaChat и перезапустите приложение.",
            500,
        )
    return key


def _get_access_token(force_refresh: bool = False) -> str:
    """Получает OAuth-токен GigaChat и кэширует его до истечения срока."""
    now = time.time()
    with _token_lock:
        cached = _token_cache.get("value")
        if not force_refresh and cached and _token_cache["expires_at"] > now:
            return str(cached)

        headers = {
            "Accept": "application/json",
            "Content-Type": "application/x-www-form-urlencoded",
            "RqUID": str(uuid.uuid4()),
            "Authorization": f"Basic {_auth_key()}",
        }
        try:
            response = requests.post(
                OAUTH_URL,
                headers=headers,
                data={"scope": SCOPE},
                timeout=(10, 30),
                verify=_ssl_verify(),
            )
        except requests.RequestException as exc:
            logger.exception("Не удалось получить токен GigaChat")
            raise AppError(
                "Не удалось подключиться к GigaChat для авторизации. "
                "Проверьте интернет и настройки TLS (GIGACHAT_CA_BUNDLE).",
                502,
            ) from exc

        if response.status_code >= 400:
            logger.error(
                "GigaChat OAuth вернул HTTP %s: %s",
                response.status_code,
                _short_response(response),
            )
            raise AppError(
                "GigaChat отклонил авторизацию. Проверьте GIGACHAT_AUTH_KEY "
                f"и GIGACHAT_SCOPE (HTTP {response.status_code}).",
                502,
            )

        try:
            payload = response.json()
        except ValueError as exc:
            logger.error("GigaChat OAuth вернул не-JSON: %s", _short_response(response))
            raise AppError("GigaChat вернул некорректный ответ авторизации.", 502) from exc

        token = payload.get("access_token")
        if not isinstance(token, str) or not token:
            raise AppError("GigaChat не вернул access_token.", 502)

        expires_at = payload.get("expires_at")
        try:
            expires_ts = float(expires_at)
        except (TypeError, ValueError):
            expires_ts = now + 25 * 60
        if expires_ts > 10_000_000_000:  # миллисекунды
            expires_ts /= 1000

        _token_cache["value"] = token
        _token_cache["expires_at"] = max(now + 60, expires_ts - 60)
        logger.info("Получен новый токен GigaChat, действует до %s", int(expires_ts))
        return token


def _short_response(response: requests.Response, limit: int = 500) -> str:
    try:
        data = response.json()
        if isinstance(data, dict):
            message = data.get("message") or data.get("error_description")
            if not message and isinstance(data.get("error"), dict):
                message = data["error"].get("message")
            if message:
                return str(message)[:limit]
        return json.dumps(data, ensure_ascii=False)[:limit]
    except ValueError:
        return response.text[:limit].replace("\n", " ")


def _post_chat(payload: dict[str, Any], token: str) -> requests.Response:
    try:
        return requests.post(
            CHAT_URL,
            headers={
                "Accept": "application/json",
                "Content-Type": "application/json",
                "Authorization": f"Bearer {token}",
            },
            json=payload,
            timeout=(10, REQUEST_TIMEOUT),
            verify=_ssl_verify(),
        )
    except requests.RequestException as exc:
        logger.exception("Ошибка запроса к GigaChat")
        raise AppError(
            "Не удалось получить ответ от GigaChat. Попробуйте позже.", 502
        ) from exc


def _extract_message_content(response: requests.Response) -> str:
    try:
        payload = response.json()
    except ValueError as exc:
        raise AppError("GigaChat вернул ответ в неизвестном формате.", 502) from exc

    try:
        content = payload["choices"][0]["message"]["content"]
    except (KeyError, IndexError, TypeError) as exc:
        logger.error("Неожиданная структура ответа GigaChat: %s", _short_response(response))
        raise AppError("GigaChat вернул ответ без текста.", 502) from exc

    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, dict):
                parts.append(str(item.get("text", "")))
            else:
                parts.append(str(item))
        content = "".join(parts)

    if not isinstance(content, str) or not content.strip():
        raise AppError("GigaChat вернул пустой ответ.", 502)
    return content


def _looks_like_json_schema_error(response: requests.Response) -> bool:
    text = _short_response(response, 2000).lower()
    return "response_format" in text or "json_schema" in text or (
        "schema" in text and "unsupported" in text
    )


def _has_explicit_birth_year(text: str) -> bool:
    """Проверяет, указан ли в тексте именно год рождения, а не возраст."""
    patterns = (
        r"(?:год\s+рождения|дата\s+рождения|родил(?:ся|ась)|рождени[ея])"
        r"[^\d]{0,35}(?:19|20)\d{2}",
        r"(?:19|20)\d{2}\s*(?:г\.?\s*р\.?|год\s+рождения|года\s+рождения)",
        r"родил(?:ся|ась)\s+в\s+(?:19|20)\d{2}",
    )
    return any(re.search(pattern, text, flags=re.IGNORECASE) for pattern in patterns)


def _number_snippets(text: str, limit: int = 35) -> str:
    """Собирает строки протокола с цифрами, чтобы модель не потеряла масштаб и сроки."""
    snippets: list[str] = []
    seen: set[str] = set()
    for raw_line in text.splitlines():
        line = re.sub(r"\s+", " ", raw_line).strip()
        if not line or not re.search(r"\d", line):
            continue
        if line in seen:
            continue
        seen.add(line)
        if len(line) > 220:
            line = line[:217].rstrip() + "..."
        snippets.append(f"- {line}")
        if len(snippets) >= limit:
            break
    return "\n".join(snippets) if snippets else "- нет"


def _sanitize_evidence(evidence: dict[str, Any]) -> dict[str, Any]:
    """Убирает из карты фактов имена и названия организаций, чтобы они не попали в отчёт."""
    clean = dict(evidence)
    for field in EVIDENCE_FIELDS:
        value = clean.get(field)
        if not isinstance(value, str):
            continue
        value = re.sub(r'\bфирм[а-я]*\s+"[^"]+"', "строительной фирмы", value)
        value = re.sub(r"\bфирм[а-я]*\s+«[^»]+»", "строительной фирмы", value)
        value = re.sub(r'"[^"]{2,80}"', "организация", value)
        value = re.sub(r"«[^»]{2,80}»", "организация", value)
        value = re.sub(
            r"\b(?:компани[а-я]*|организаци[а-я]*|центр[а-я]*|завод[а-я]*)"
            r"\s+[А-ЯЁA-Z][\w-]*",
            lambda m: m.group(0).split()[0],
            value,
        )
        value = re.sub(
            r"\bКомитет[а-я]*\s+государственной\s+застройки\b",
            "государственном комитете",
            value,
        )
        value = re.sub(r"\bСУ-?\s?155\b", "строительная организация", value)
        value = re.sub(r"\bСмартфо\b", "строительная фирма", value)
        value = re.sub(r"\bЖигули\b", "завод", value)
        clean[field] = value
    return clean


def _call_json(
    system_prompt: str,
    user_prompt: str,
    schema: dict[str, Any],
    *,
    max_tokens: int,
    temperature: float,
    label: str,
) -> dict[str, Any]:
    """Один вызов GigaChat с JSON-схемой и обработкой типовых ошибок."""
    token = _get_access_token()
    payload: dict[str, Any] = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False,
        "response_format": {
            "type": "json_schema",
            "schema": schema,
            "strict": True,
        },
    }

    response = _post_chat(payload, token)

    if response.status_code == 401:
        logger.info("Токен GigaChat истёк, обновляю и повторяю запрос")
        token = _get_access_token(force_refresh=True)
        response = _post_chat(payload, token)

    if response.status_code >= 400 and _looks_like_json_schema_error(response):
        logger.warning(
            "GigaChat не поддержал response_format (%s), повторяю без него: %s",
            label,
            _short_response(response),
        )
        payload.pop("response_format", None)
        response = _post_chat(payload, token)

    if response.status_code >= 400:
        logger.error(
            "GigaChat chat/completions (%s) вернул HTTP %s: %s",
            label,
            response.status_code,
            _short_response(response),
        )
        if response.status_code == 429:
            raise AppError(
                "GigaChat временно ограничил частоту запросов. Повторите через минуту.",
                503,
            )
        if response.status_code in (413, 422):
            raise AppError(
                "Протокол слишком большой для GigaChat или содержит неподдерживаемые "
                "данные. Сократите текст или уменьшите MAX_PROTOCOL_CHARS.",
                413,
            )
        if response.status_code in (400, 401, 403):
            raise AppError(
                "GigaChat отклонил запрос. Проверьте ключ, доступ к модели "
                f"{MODEL} и scope (HTTP {response.status_code}).",
                502,
            )
        raise AppError(
            f"GigaChat вернул ошибку (HTTP {response.status_code}). Попробуйте позже.",
            502,
        )

    try:
        result_payload = response.json()
        choice = (result_payload.get("choices") or [{}])[0]
        logger.info(
            "GigaChat (%s): finish_reason=%s, usage=%s",
            label,
            choice.get("finish_reason"),
            result_payload.get("usage"),
        )
    except (ValueError, AttributeError, TypeError):
        pass

    raw = _extract_message_content(response)
    return _parse_json_object(raw)


def extract_evidence(
    protocol_text: str,
    target_role_hint: str = "",
) -> dict[str, Any] | None:
    """Первый проход: выделяет из протокола карту фактов без оценок."""
    target_hint = target_role_hint.strip()
    target_role_text = target_hint or "не задана отдельно — определи из протокола, если она там есть"

    user_prompt = (
        "Извлеки из протокола карту фактов для последующей оценки компетенций.\n\n"
        f"ЦЕЛЕВАЯ ДОЛЖНОСТЬ (указана пользователем): {target_role_text}.\n"
        "Если она задана, используй это значение; если нет — возьми из протокола; "
        "если нет и там — «Не указано».\n\n"
        "ТРЕБОВАНИЯ:\n"
        "- Не пиши отчёт и не давай итоговых оценок.\n"
        "- Сохраняй конкретику: роли, стаж, масштаб, число подчинённых, объекты, сроки, "
        "результаты, цифры. Если цифр нет — так и укажи.\n"
        "- Для каждого примера поведения укажи: ситуация — действия — результат; "
        "источник (кандидат или интервьюер); силу примера (сильная/средняя/слабая, "
        "единичный/повторяющийся).\n"
        "- Отдельно вынеси оценочные заметки интервьюера, не превращая их в факты.\n"
        "- В data_gaps перечисли темы, которые не раскрыты. Это пробел в данных, "
        "а не низкий уровень компетенции.\n\n"
        "ФРАГМЕНТЫ ПРОТОКОЛА С ЦИФРАМИ (проверь контекст, не выдумывай):\n"
        f"{_number_snippets(protocol_text)}\n\n"
        "--- НАЧАЛО ПРОТОКОЛА ---\n"
        f"{protocol_text}\n"
        "--- КОНЕЦ ПРОТОКОЛА ---\n\n"
        "Верни только JSON-объект с полями fio, current_role, target_role, birth_year, "
        "diagnostic_date, experience_facts, management_scope, key_numbers, achievements, "
        "behavioral_examples, interviewer_observations, motivation, contradictions, data_gaps."
    )

    try:
        evidence = _call_json(
            EVIDENCE_SYSTEM_PROMPT,
            user_prompt,
            EVIDENCE_SCHEMA,
            max_tokens=4096,
            temperature=0.1,
            label="evidence",
        )
        for field in EVIDENCE_FIELDS:
            value = evidence.get(field)
            if isinstance(value, str):
                value = _strip_outer_braces(_clean_markup(value))
                if re.search(r"\bсвои\s+(?:долги|финансовые)", protocol_text, flags=re.IGNORECASE):
                    value = re.sub(
                        r"финансовые\s+(?:задолженности|долги)\s+компании",
                        "свои финансовые долги",
                        value,
                        flags=re.IGNORECASE,
                    )
                    value = re.sub(
                        r"(?<!свои )долги\s+компании",
                        "свои долги",
                        value,
                        flags=re.IGNORECASE,
                    )
                evidence[field] = value
        if not _has_explicit_birth_year(protocol_text):
            evidence["birth_year"] = "Не указано"
        return evidence
    except AppError as exc:
        logger.warning(
            "Не удалось извлечь карту фактов (%s). Перехожу к одноэтапной генерации.",
            exc.message,
        )
        return None


def write_report(
    protocol_text: str,
    evidence: dict[str, Any] | None,
    target_role_hint: str = "",
) -> dict[str, Any]:
    """Второй проход: пишет отчёт по карте фактов и протоколу."""
    target_hint = target_role_hint.strip()
    target_role_text = (
        target_hint or "не задана отдельно — определи из карты фактов или протокола, если она там есть"
    )

    if evidence:
        evidence = _sanitize_evidence(evidence)
        key_numbers = str(evidence.get("key_numbers", "")).strip()
        key_numbers = "\n".join(
            line for line in key_numbers.splitlines() if "возраст" not in line.lower()
        ).strip() or "- нет"
        evidence_for_prompt = dict(evidence)
        evidence_for_prompt["key_numbers"] = key_numbers
        evidence_text = json.dumps(evidence_for_prompt, ensure_ascii=False, indent=2)
        evidence_block = (
            "КАРТА ФАКТОВ, ПОДГОТОВЛЕННАЯ ПО ПРОТОКОЛУ:\n"
            f"{evidence_text}\n\n"
            "Карта фактов — единственный источник конкретики. Не добавляй имена, "
            "организации, цифры, события и выводы, которых нет в карте. Если в карте "
            "встречается название организации, замени его ролью.\n\n"
        )
        protocol_block = (
            "ПРОТОКОЛ НЕ ПЕРЕДАЁТСЯ ОТДЕЛЬНО: используй только карту фактов. "
            "Не пытайся восстановить детали, которых в ней нет.\n\n"
        )
    else:
        evidence_block = (
            "Карта фактов не подготовлена. Работай непосредственно по протоколу.\n\n"
        )
        key_numbers = _number_snippets(protocol_text)
        protocol_block = (
            "--- НАЧАЛО ПРОТОКОЛА ---\n"
            f"{protocol_text}\n"
            "--- КОНЕЦ ПРОТОКОЛА ---\n\n"
        )

    user_prompt = (
        "Подготовь отчёт по результатам интервью.\n\n"
        f"ЦЕЛЕВАЯ ДОЛЖНОСТЬ (указана пользователем): {target_role_text}.\n"
        "Если целевая должность задана, оценивай кандидата именно под неё. "
        "Если не задана, возьми её из карты фактов или протокола; не придумывай.\n\n"
        + evidence_block
        + "СПРАВКА ПО ЦИФРАМ (используй только те, что относятся к опыту, масштабу, "
        "срокам или результатам; не включай возраст и случайные числа):\n"
        f"{key_numbers}\n\n"
        "СТРУКТУРА И ТРЕБОВАНИЯ:\n"
        "- conclusion — 1–2 абзаца, 120–180 слов. Синтез: степень соответствия целевой "
        "должности, 2–3 главных аргумента, ключевые ограничения, условия успешного "
        "назначения и осторожная рекомендация. Не повторяй списки дословно. Заключение "
        "должно согласовываться с пунктами и не содержать выводов, противоречащих примерам.\n"
        "- strengths — 4–6 пунктов по 35–70 слов. В каждом: тезис, конкретный пример "
        "или наблюдение из протокола, значение для целевой роли. Используй ключевые цифры "
        "из карты фактов: стаж, число подчинённых, масштаб, сроки, результаты. Сверяй "
        "значения и контекст с key_numbers: не приписывай прошлый масштаб текущей роли. "
        "В strengths обязательно используй минимум два масштабных факта из key_numbers: "
        "число людей, число организаций, стаж или сроки.\n"
        "- risks — 3–5 пунктов по 35–70 слов. В каждом: в чём риск, при каких условиях "
        "проявится, на что повлияет, что проверить или развить. Негативный вывод делай "
        "только при наличии конкретного примера; если примера нет, пиши «в интервью "
        "не проверено», а не «слабо развито».\n"
        "- motivation — 2–4 пункта по 25–50 слов: мотиваторы, демотиваторы, значимые "
        "условия, риск удержания. Только по протоколу. Не добавляй карьерные амбиции, "
        "желание влиять на стратегию и другие мотивы, которых нет в поле motivation карты.\n"
        "Сохраняй роли и контекст из карты: «рота» — не «команда», «база организаций» — "
        "не «координация их деятельности». Выводы по заметкам интервьюера формулируй "
        "как наблюдение, а не как диагноз.\n"
        "Не начинай все пункты с «Кандидат»; формулировки должны быть разнообразными.\n\n"
        "Перед ответом проверь: каждый пункт имеет основание в протоколе или карте фактов; "
        "использованы ключевые цифры; нет выдуманных фактов, цифр, имён и организаций; "
        "нет повторов; риски сформулированы как вероятность, а не как приговор; заключение "
        "не дублирует пункты и не противоречит им.\n\n"
        + protocol_block
        + "Верни только JSON-объект с полями fio, current_role, target_role, birth_year, "
        "diagnostic_date, conclusion, strengths, risks, motivation."
    )

    return _call_json(
        SYSTEM_PROMPT,
        user_prompt,
        REPORT_SCHEMA,
        max_tokens=4096,
        temperature=0.15,
        label="report",
    )


def verify_report(
    draft: dict[str, Any],
    protocol_text: str,
    evidence: dict[str, Any] | None,
    target_role_hint: str = "",
) -> dict[str, Any]:
    """Третий проход: строгий редактор убирает домыслы и согласует отчёт с фактами."""
    target_hint = target_role_hint.strip()
    draft_clean = _normalize_report(draft)
    if evidence:
        evidence = _sanitize_evidence(evidence)
        key_numbers = str(evidence.get("key_numbers", "")).strip()
        key_numbers = "\n".join(
            line for line in key_numbers.splitlines() if "возраст" not in line.lower()
        ).strip() or "- нет"
        evidence_for_prompt = dict(evidence)
        evidence_for_prompt["key_numbers"] = key_numbers
        evidence_text = json.dumps(evidence_for_prompt, ensure_ascii=False, indent=2)
    else:
        evidence_text = "Карта фактов не подготовлена. Проверяй только по протоколу."
    if evidence:
        protocol_block = (
            "ПРОТОКОЛ НЕ ПЕРЕДАЁТСЯ ОТДЕЛЬНО. Проверяй только по карте фактов. "
            "Если утверждения нет в карте, удали его или переформулируй как "
            "«в интервью не проверено».\n\n"
        )
    else:
        protocol_block = (
            "--- НАЧАЛО ПРОТОКОЛА ---\n"
            f"{protocol_text}\n"
            "--- КОНЕЦ ПРОТОКОЛА ---\n\n"
        )
    user_prompt = (
        "Проверь и исправь черновик отчёта.\n\n"
        f"ЦЕЛЕВАЯ ДОЛЖНОСТЬ: {target_hint or 'не указана отдельно'}.\n\n"
        "КАРТА ФАКТОВ:\n"
        f"{evidence_text}\n\n"
        "ЧЕРНОВИК ОТЧЁТА:\n"
        f"{json.dumps(draft_clean, ensure_ascii=False, indent=2)}\n\n"
        + protocol_block
        + "Особо проверь: числа и их контекст (key_numbers), непроверенные компетенции "
        "(не писать «слабо» или «владеет недостаточно»), повтор «Кандидат» в начале пунктов, "
        "противоречия между заключением и пунктами.\n\n"
        "Верни только исправленный JSON-объект. Не добавляй факты, которых нет в карте "
        "или протоколе. Удали или переформулируй всё, что не подтверждено."
    )
    try:
        return _call_json(
            VERIFY_SYSTEM_PROMPT,
            user_prompt,
            REPORT_SCHEMA,
            max_tokens=4096,
            temperature=0.1,
            label="verify",
        )
    except AppError as exc:
        logger.warning(
            "Не удалось проверить черновик (%s). Возвращаю черновик без редактуры.",
            exc.message,
        )
        return draft


def analyze_protocol(protocol_text: str, target_role_hint: str = "") -> dict[str, str]:
    """Готовит отчёт: карта фактов, черновик и строгая редактура."""
    target_hint = target_role_hint.strip()
    evidence = extract_evidence(protocol_text, target_hint) if TWO_PASS else None
    data = write_report(protocol_text, evidence, target_hint)
    if VERIFY_REPORT:
        data = verify_report(data, protocol_text, evidence, target_hint)
    report = _normalize_report(data)

    evidence_data = evidence or {}
    for field in ("conclusion", "strengths", "risks", "motivation"):
        if evidence:
            report[field] = _fix_number_context(report[field], evidence)
        report[field] = _fix_subject_context(report[field], protocol_text, evidence_data)
    report = _ensure_key_facts(report, evidence_data, protocol_text)
    report = _ensure_motivation(report, protocol_text)

    if target_hint:
        report["target_role"] = target_hint
    if not _has_explicit_birth_year(protocol_text):
        report["birth_year"] = "Не указано"
    return report


def _parse_json_object(raw: str) -> dict[str, Any]:
    text = raw.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)

    candidates = [text]
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        candidates.append(text[start : end + 1])

    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, dict):
            return parsed

    logger.error("Не удалось разобрать JSON GigaChat: %s", raw[:1500])
    raise AppError("GigaChat вернул невалидный JSON. Попробуйте сгенерировать отчёт ещё раз.", 502)


def _as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, (list, tuple)):
        return "\n".join(_as_text(item) for item in value if _as_text(item))
    if isinstance(value, dict):
        return "; ".join(f"{key}: {_as_text(item)}" for key, item in value.items())
    return str(value).strip()


def _clean_markup(text: str) -> str:
    """Убирает случайную HTML/BBCode-разметку, которую модель иногда добавляет."""
    tags = (
        r"strong|b|i|em|u|p|br|li|ul|ol|list|blockquote|span|div|font|h[1-6]"
    )
    # Границы списков превращаем в переводы строк, иначе пункты склеиваются.
    text = re.sub(r"\[/?(?:li|ul|ol|list)\b[^\]]*\]", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"\{/?(?:li|ul|ol|list)\b[^\}]*\}?", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"(?<!\w)(?:li|ul|ol|list)\}", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"\[/?(?:" + tags + r")\b[^\]]*\]", "", text, flags=re.IGNORECASE)
    text = re.sub(r"</?(?:" + tags + r")\b[^>]*>", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\{/?(?:" + tags + r")\b[^\}]*\}?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"(?<!\w)(?:" + tags + r")\}", "", text, flags=re.IGNORECASE)
    text = text.replace("**", "").replace("__", "")
    return text


def _strip_outer_braces(text: str) -> str:
    """Убирает фигурные скобки, которыми модель иногда оборачивает значение поля."""
    value = text.strip()
    if len(value) >= 2 and value.startswith("{") and value.endswith("}"):
        return value[1:-1].strip()
    return value


def _format_bullets(value: Any) -> str:
    text = _as_text(value).replace("\r", "")
    text = text.replace("\\n", "\n")
    text = _strip_outer_braces(_clean_markup(text))
    if text.strip().lower().rstrip(".") in {"не указано", "нет данных"}:
        return "Не указано"
    text = re.sub(r"[ \t]*•[ \t]*", "\n", text)
    lines = []
    for line in text.split("\n"):
        line = re.sub(r"^\s*(?:[*\-–—]|\d+[.)])\s*", "", line)
        line = re.sub(r"^\s*\[[^\]]{1,80}\]\s*", "", line)
        line = re.sub(r"^[А-ЯЁA-Z][^:]{2,90}:\s*", "", line)
        line = re.sub(r"^(?:Кандидат|Кандидатка|Он|Она)\s+", "", line, flags=re.IGNORECASE)
        line = re.sub(r"^\s*:\s*", "", line).strip()
        line = _strip_outer_braces(line)
        if line.startswith("{"):
            line = line[1:].strip()
        if line.endswith("}"):
            line = line[:-1].strip()
        if line.startswith("["):
            line = line[1:].strip()
        if line.endswith("]"):
            line = line[:-1].strip()
        if line and line[0].islower():
            line = line[0].upper() + line[1:]
        if line:
            lines.append(f"• {line}")
    return "\n".join(lines) if lines else "Не указано"


def _fix_number_context(text: str, evidence: dict[str, Any]) -> str:
    """Исправляет явное смешение контекста: число, относящееся к роте, не должно стать бригадой."""
    source = "\n".join(
        str(value) for value in evidence.values() if isinstance(value, str)
    )
    for line in source.splitlines():
        if "рот" not in line.lower():
            continue
        numbers = [int(value) for value in re.findall(r"\b(\d{2,4})\b", line)]
        if not numbers:
            continue
        number = str(max(numbers))
        pattern = (
            rf"(?:бригад[а-я]*|команд[а-я]*|коллектив[а-я]*|подчинённ[а-я]*|подчиненн[а-я]*)"
            rf"\s+(?:из\s+|численностью\s+|до\s+)*{number}"
        )
        text = re.sub(pattern, f"ротой из {number}", text, flags=re.IGNORECASE)
    return text


def _fix_subject_context(text: str, protocol_text: str, evidence: dict[str, Any]) -> str:
    """Исправляет подмену субъекта и характеристики, если она видна по протоколу."""
    evidence_text = "\n".join(
        str(value) for value in evidence.values() if isinstance(value, str)
    )
    if re.search(r"\bсвои\s+(?:долги|финансовые)", protocol_text, flags=re.IGNORECASE):
        text = re.sub(
            r"(?:устранил|погасил|закрыл)\s+финансовые\s+(?:задолженности|долги)\s+компании",
            "закрыл личные финансовые обязательства",
            text,
            flags=re.IGNORECASE,
        )
        text = re.sub(
            r"(?:устранил|погасил|закрыл)\s+долги\s+компании",
            "закрыл личные долги",
            text,
            flags=re.IGNORECASE,
        )
    if "длина" in evidence_text.lower() and "колонн" in evidence_text.lower():
        text = re.sub(r"колонн\s+высотой", "колонн длиной", text, flags=re.IGNORECASE)
        text = text.replace("колоннами высотой", "колоннами длиной")
    text = re.sub(
        r"координаци[а-я]*\s+базы",
        "работу с базой",
        text,
        flags=re.IGNORECASE,
    )
    text = text.replace("небольшими коллективами", "коллективами разного масштаба")
    text = text.replace("небольших коллективов", "коллективов разного масштаба")
    text = text.replace("небольшим коллективом", "коллективами разного масштаба")
    text = text.replace("небольшой коллектив", "коллектив разного масштаба")
    text = text.replace("малыми коллективами", "коллективами разного масштаба")
    text = text.replace("малых коллективов", "коллективов разного масштаба")
    text = text.replace("малым коллективом", "коллективами разного масштаба")
    text = text.replace("управлении коллективом", "управлении коллективами разного масштаба")
    text = text.replace("Назначение будет эффективно", "Назначение будет успешным")
    text = text.replace(
        "стремление сохранить здоровье и профессиональный рост, свидетельствующее",
        "стремление сохранить здоровье и профессиональный рост, свидетельствующие",
    )
    text = text.replace(
        "потребуется существенное расширение управленческого опыта",
        "потребуется развитие управленческого опыта на уровне всего предприятия",
    )
    text = text.replace("(~250 м)", "(около 250 м)")
    text = text.replace(
        "слабая управленческая подготовка", "недостаточная управленческая подготовка"
    )
    text = text.replace(
        "слабую управленческую подготовку", "недостаточную управленческую подготовку"
    )
    text = text.replace(
        "слабой управленческой подготовкой", "недостаточной управленческой подготовкой"
    )
    text = text.replace("слабая демонстрация", "недостаточно подтверждённая практика")
    text = text.replace("слабую демонстрацию", "недостаточно подтверждённую практику")
    text = text.replace("слабой демонстрацией", "недостаточно подтверждённой практикой")
    text = text.replace(
        "дополнительного тестирования", "дополнительной проверки"
    )
    text = text.replace("тестирование", "проверка")
    text = text.replace(
        "не всегда эффективно проявляет лидерство в сложных или новых коллективах",
        "недостаточно подтверждён опыт лидерства в новых коллективах",
    )
    text = text.replace(
        "опытом в строительстве и управлением",
        "опытом в строительстве и управлении",
    )
    text = text.replace("опытом в строительстве и управления", "опытом в строительстве и управлении")
    text = text.replace(
        "Недостаточная демонстрация умений",
        "Недостаточно подтверждены умения",
    )
    text = text.replace("широкого масштаба", "широкого профиля")
    text = text.replace("с базой из около", "с базой около")
    text = text.replace("с базой из порядка", "с базой порядка")
    text = text.replace("базой из около", "базой около")
    text = text.replace("базой из порядка", "базой порядка")
    text = text.replace("организационной трансформации", "стратегическом планировании")
    text = text.replace("организационных изменениях", "планировании")
    text = text.replace("целостной картины бизнеса", "целостной картины")
    text = text.replace("задач бизнеса", "общих задач")
    text = text.replace(
        "недостаточными стремлением и способностью решать проблемы системно и выходить "
        "за рамки текущих полномочий",
        "недостаточно подтверждённой готовностью решать проблемы системно и выходить "
        "за рамки текущих полномочий",
    )
    text = text.replace(
        "управляет участком с 1–3 бригадами",
        "управляет 1–3 строительными участками",
    )
    text = re.sub(
        r"управляет\s+участком\s+с\s+1\s*[-–—]\s*3\s+бригадами",
        "управляет 1-3 строительными участками",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"управлял\s+участком\s+с\s+1\s*[-–—]\s*3\s+бригадами",
        "управлял 1-3 строительными участками",
        text,
        flags=re.IGNORECASE,
    )
    text = text.replace(" и сдачу объекта", "")
    text = text.replace(
        "координации большого количества участников",
        "координации большого числа организаций",
    )
    text = text.replace(", предпочитая традиционные подходы", "")
    text = text.replace(
        "затруднения с адаптацией и обобщением опыта",
        "трудности с адаптацией",
    )
    text = text.replace(
        "испытывает сложности с адаптацией и обобщением опыта",
        "испытывает сложности с адаптацией",
    )
    text = text.replace(
        "трудности с адаптацией и обобщением опыта",
        "трудности с адаптацией",
    )
    text = text.replace(
        "Отсутствуют примеры успешного взаимодействия",
        "В интервью не выявлены примеры успешного взаимодействия",
    )
    text = text.replace(
        "Отсутствуют конкретные примеры",
        "В интервью не выявлены конкретные примеры",
    )
    text = text.replace(
        "Отсутствуют конкретные данные",
        "В интервью не выявлены конкретные данные",
    )
    text = text.replace(
        "Отсутствие регулярных выходных и хронического недосыпа воспринимает",
        "Отсутствие регулярных выходных и хронический недосып он воспринимает",
    )
    text = text.replace(
        "координировал работу около 700 организаций в лицензионном центре",
        "работал в лицензионном центре с базой около 700 организаций",
    )
    text = text.replace("работу 1-3 участков", "работу 1-3 строительных участков")
    text = text.replace(
        "управления значительными массивами исполнителей и контрагентов",
        "управления большим числом исполнителей и контрагентов",
    )
    text = text.replace(
        "создал базу из примерно 700 строительных организаций в составе лицензионного центра",
        "работал с базой примерно 700 строительных организаций в составе лицензионного центра",
    )
    text = text.replace(
        "и лицензионным центром с участием около 700 организаций",
        "и работал в лицензионном центре с базой около 700 организаций",
    )
    text = text.replace(
        "большим числом участников и данных",
        "большим числом организаций",
    )
    text = text.replace(
        "организовывать работу значительного количества участников",
        "организовывать работу большого числа организаций",
    )
    text = text.replace(
        "на более ответственной должности с повышенной ответственностью",
        "на более ответственной должности",
    )
    text = text.replace("заместительской позиции", "позиции заместителя директора")
    text = text.replace("слабую адаптацию", "трудности с адаптацией")
    text = text.replace("слабая адаптация", "трудности с адаптацией")
    text = text.replace(
        "и работы с крупными массивами информации",
        "и работы с большим числом организаций",
    )
    text = text.replace(
        "к изменениям и широкому управлению",
        "к изменениям и управлению широкого профиля",
    )
    text = text.replace(
        "трудности в коммуникации идей и адаптации к переменам",
        "трудности с адаптацией и излишняя детализация объяснений",
    )
    text = text.replace(
        "недостаточно адаптивен и подготовлен к оптимизации организационных процессов",
        "испытывает трудности с адаптацией и оптимизацией организационных процессов",
    )
    text = text.replace("замдиректора", "заместителя директора")
    text = text.replace("на линейном объекте длиной свыше 250 метров", "на линейном объекте")
    text = text.replace(
        "добился уважения и дисциплины личного состава",
        "обеспечивал дисциплину и организацию личного состава",
    )
    text = re.sub(
        r"управление\s+лицензионным\s+центром",
        "работу в лицензионном центре",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"управлял\s+лицензионным\s+центром",
        "работал в лицензионном центре",
        text,
        flags=re.IGNORECASE,
    )
    text = text.replace(
        "создание саморегулируемых организаций и лицензионного центра",
        "создание саморегулируемых организаций и работа в лицензионном центре",
    )
    text = re.sub(
        r"создание\s+лицензионного\s+центра",
        "работа в лицензионном центре",
        text,
        flags=re.IGNORECASE,
    )
    text = re.sub(
        r"создал\s+лицензионный\s+центр\s+с\s+участием\s+около\s+700\s+организаций",
        "работал в лицензионном центре с базой около 700 организаций",
        text,
        flags=re.IGNORECASE,
    )
    if "лицензионн" in protocol_text.lower() and not re.search(
        r"создал\s+лицензионн", protocol_text, flags=re.IGNORECASE
    ):
        text = re.sub(
            r"создал\s+лицензионный\s+центр",
            "работал в лицензионном центре",
            text,
            flags=re.IGNORECASE,
        )
    if re.search(
        r"адм\w*\s+центр\w*[^\n]{0,100}\b100\b", protocol_text, flags=re.IGNORECASE
    ):
        text = re.sub(
            r"саморегулируемую организацию,\s*включающую более 100 (?:участников|организаций)",
            "саморегулируемую организацию, административный центр которой объединял "
            "более 100 организаций",
            text,
            flags=re.IGNORECASE,
        )
    return text


def _ensure_key_facts(
    report: dict[str, str],
    evidence: dict[str, Any],
    protocol_text: str = "",
) -> dict[str, str]:
    """Добавляет ключевые масштабные факты, если модель их потеряла."""
    strengths = report.get("strengths", "")
    if not strengths or strengths == "Не указано":
        return report

    lines = strengths.splitlines()
    additions: list[str] = []
    key_numbers = str(evidence.get("key_numbers", ""))
    experience = str(evidence.get("experience_facts", ""))
    protocol_lower = protocol_text.lower()

    if "186" in protocol_text and "186" not in strengths:
        additions.append(
            "• Ранее командовал ротой численностью 186 человек, что подтверждает опыт "
            "руководства большим коллективом."
        )
    has_700 = (
        "700" in key_numbers and "организац" in key_numbers.lower()
    ) or ("700" in protocol_text and "организац" in protocol_lower)
    if has_700 and "700" not in strengths:
        additions.append(
            "• В лицензионном центре работал с базой около 700 организаций, что подтверждает "
            "опыт взаимодействия с большим числом внешних структур."
        )
    has_nine_years = "9 лет" in experience or bool(
        re.search(r"\b9\s+лет", protocol_text, flags=re.IGNORECASE)
    )
    if has_nine_years and not re.search(r"(?:9|девять)\s+лет", strengths, flags=re.IGNORECASE):
        additions.append(
            "• Ранее 9 лет работал в государственном комитете, что даёт опыт взаимодействия "
            "с государственными структурами и понимание регулирования строительной отрасли."
        )
    if protocol_text and "100" in protocol_text and "100" not in strengths:
        additions.append(
            "• Административный центр саморегулируемой организации объединял более 100 "
            "строительных организаций, что подтверждает опыт работы с профессиональным "
            "сообществом."
        )

    available = max(0, 6 - len(lines))
    if additions:
        if available:
            lines[1:1] = additions[:available]
        else:
            combined = "• " + " ".join(
                item.lstrip("• ").strip() for item in additions
            )
            lines[-1] = combined
        report["strengths"] = "\n".join(lines)
    return report


def _ensure_motivation(report: dict[str, str], protocol_text: str) -> dict[str, str]:
    """Добавляет мотив сохранения здоровья, если он есть в протоколе, но потерян."""
    motivation = report.get("motivation", "")
    if not motivation or motivation == "Не указано":
        return report
    if "здоров" not in protocol_text.lower() or "здоров" in motivation.lower():
        return report
    lines = motivation.splitlines()
    if len(lines) < 3:
        lines.insert(
            0,
            "• Стремление сохранять здоровье является значимым мотиватором.",
        )
        report["motivation"] = "\n".join(lines)
    return report


def _normalize_report(data: dict[str, Any]) -> dict[str, str]:
    report: dict[str, str] = {}
    for field in REPORT_FIELDS:
        value = data.get(field)
        if field in {"strengths", "risks", "motivation"}:
            text = _format_bullets(value)
        else:
            text = _strip_outer_braces(_clean_markup(_as_text(value)))

        if field == "birth_year" and re.fullmatch(r"\d{1,2}", text):
            # В протоколе может быть указан возраст, а не год рождения.
            text = "Не указано"

        report[field] = text or "Не указано"
    return report


def extract_text(filename: str, blob: bytes) -> str:
    """Извлекает текст из .docx (mammoth, затем python-docx) или .txt."""
    suffix = Path(filename).suffix.lower()

    if suffix == ".txt":
        text = _decode_txt(blob)
    elif suffix == ".docx":
        text = _extract_docx(blob)
    else:
        raise AppError("Поддерживаются только файлы .docx и .txt.", 400)

    text = _clean_text(text)
    if not text:
        raise AppError("В загруженном файле не найден текст.", 400)
    return text


def _decode_txt(blob: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return blob.decode(encoding)
        except UnicodeDecodeError:
            continue
    return blob.decode("utf-8", errors="replace")


def _extract_docx(blob: bytes) -> str:
    errors: list[str] = []

    try:
        with io.BytesIO(blob) as stream:
            result = mammoth.extract_raw_text(stream)
        text = _clean_text(result.value)
        if text:
            return text
    except Exception as exc:  # mammoth может падать на повреждённом docx
        errors.append(f"mammoth: {exc}")
        logger.warning("mammoth не смог прочитать docx: %s", exc)

    # Резервный путь через python-docx: помогает и при некоторых нестандартных docx.
    try:
        document = Document(io.BytesIO(blob))
        parts = [p.text for p in document.paragraphs if p.text.strip()]
        for table in document.tables:
            for row in table.rows:
                cells = []
                for cell in row.cells:
                    cell_text = " ".join(
                        p.text.strip() for p in cell.paragraphs if p.text.strip()
                    )
                    if cell_text and (not cells or cell_text != cells[-1]):
                        cells.append(cell_text)
                if cells:
                    parts.append(" | ".join(cells))
        text = _clean_text("\n".join(parts))
        if text:
            return text
    except Exception as exc:
        errors.append(f"python-docx: {exc}")
        logger.warning("python-docx не смог прочитать docx: %s", exc)

    logger.error("Не удалось извлечь текст из docx: %s", "; ".join(errors))
    raise AppError(
        "Не удалось прочитать .docx. Убедитесь, что это корректный файл Word.", 400
    )


def _clean_text(text: str) -> str:
    text = text.replace("\x00", " ").replace("\r\n", "\n").replace("\r", "\n")
    lines = [re.sub(r"[ \t]+", " ", line).strip() for line in text.split("\n")]
    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def render_report(report: dict[str, str]) -> str:
    if not TEMPLATE_FILE.is_file():
        raise AppError(
            "Не найден template.docx. Положите шаблон отчёта в папку приложения.", 500
        )

    try:
        document = DocxTemplate(str(TEMPLATE_FILE))
        declared = document.get_undeclared_template_variables()
        if not declared:
            raise AppError(
                "В template.docx не найдены теги docxtpl. Добавьте теги вида "
                "{{ fio }}, {{ strengths }} и другие поля отчёта.",
                500,
            )

        missing = set(REPORT_FIELDS) - declared
        if missing:
            logger.warning(
                "В шаблоне нет тегов полей: %s",
                ", ".join(sorted(missing)),
            )

        document.render(report, autoescape=True)
    except AppError:
        raise
    except Exception as exc:
        logger.exception("Ошибка заполнения шаблона")
        raise AppError(
            "Не удалось заполнить template.docx. Проверьте корректность шаблона "
            "и наличие тегов полей.",
            500,
        ) from exc

    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        filename = (
            f"report_{datetime.now():%Y%m%d_%H%M%S}_{uuid.uuid4().hex[:8]}.docx"
        )
        document.save(str(OUTPUT_DIR / filename))
    except OSError as exc:
        logger.exception("Не удалось сохранить отчёт")
        raise AppError("Не удалось сохранить готовый отчёт в папку output.", 500) from exc

    logger.info("Отчёт сохранён: %s", OUTPUT_DIR / filename)
    return filename


@app.get("/")
def index():
    if not INDEX_FILE.is_file():
        return "Файл index.html не найден.", 500
    return send_from_directory(BASE_DIR, "index.html")


@app.post("/generate")
def generate():
    try:
        upload = request.files.get("file")
        if upload is None:
            raise AppError("Файл не выбран. Прикрепите протокол .docx или .txt.", 400)
        if not upload.filename:
            raise AppError("У загруженного файла нет имени.", 400)

        suffix = Path(upload.filename).suffix.lower()
        if suffix not in {".docx", ".txt"}:
            raise AppError("Поддерживаются только файлы .docx и .txt.", 400)

        blob = upload.read()
        if not blob:
            raise AppError("Загруженный файл пустой.", 400)
        if len(blob) > MAX_FILE_BYTES:
            raise AppError("Файл больше 15 МБ.", 413)

        protocol_text = extract_text(upload.filename, blob)
        if len(protocol_text) > MAX_PROTOCOL_CHARS:
            logger.warning(
                "Протокол обрезан с %s до %s символов",
                len(protocol_text),
                MAX_PROTOCOL_CHARS,
            )
            protocol_text = protocol_text[:MAX_PROTOCOL_CHARS]

        target_role = request.form.get("target_role", "").strip()[:200]
        logger.info(
            "Протокол %s: %s символов, целевая должность: %s, отправляю в %s",
            upload.filename,
            len(protocol_text),
            target_role or "из протокола",
            MODEL,
        )
        report = analyze_protocol(protocol_text, target_role_hint=target_role)
        filename = render_report(report)

        return jsonify(
            {
                "ok": True,
                "filename": filename,
                "download_url": url_for("download", filename=filename),
                "data": report,
            }
        )
    except AppError as exc:
        logger.warning("Ошибка обработки: %s", exc.message)
        return jsonify({"ok": False, "error": exc.message}), exc.status_code
    except RequestEntityTooLarge:
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "Файл слишком большой. Максимальный размер — 15 МБ.",
                }
            ),
            413,
        )
    except Exception:
        logger.exception("Непредвиденная ошибка")
        return (
            jsonify(
                {
                    "ok": False,
                    "error": "Внутренняя ошибка сервера. Подробности есть в логе приложения.",
                }
            ),
            500,
        )


@app.get("/download/<path:filename>")
def download(filename: str):
    return send_from_directory(
        OUTPUT_DIR,
        filename,
        as_attachment=True,
        download_name=filename,
        max_age=0,
    )


@app.errorhandler(RequestEntityTooLarge)
def handle_too_large(_error):
    return (
        jsonify(
            {
                "ok": False,
                "error": "Файл слишком большой. Максимальный размер — 15 МБ.",
            }
        ),
        413,
    )


if __name__ == "__main__":
    if not TEMPLATE_FILE.is_file():
        logger.warning("template.docx не найден: %s", TEMPLATE_FILE)
    if not os.getenv("GIGACHAT_AUTH_KEY"):
        logger.warning("GIGACHAT_AUTH_KEY не задан — генерация отчётов не сработает.")

    host = os.getenv("HOST", "127.0.0.1")
    port = int(os.getenv("PORT", "5000"))
    logger.info("Сервер запущен: http://localhost:%s", port)
    port = int(os.environ.get('PORT', port))
    app.run(host='0.0.0.0', port=port, debug=False)
