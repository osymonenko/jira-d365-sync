"""Раскладка недостающих часов по строкам задач до 8 ч в день.

Чистая логика: ни openpyxl, ни сети, ни системных часов. Всё считается в
четвертях часа (int), чтобы не ловить ошибки float — но выдаём только кратное
получасу (STEP_Q), потому что дробные 0.25/0.75 в табеле неудобны.

Договорённости:
  * регулярные задачи (Internal Daily meeting и прочие из schedule) не трогаем —
    у них свои предписанные часы;
  * добиваем остаток строками из Jira и QA-заглушками;
  * обычная задача целиком укладывается в ОДИН день: никаких «1.5 ч в среду и
    ещё 0.5 ч в четверг» — в табеле такой перенос читается как незаконченная
    работа, хотя это просто артефакт упаковки;
  * «размазываемые» задачи (Bug verification, Investigation issue) — исключение:
    они идут кусками по 0.5 ч через всю неделю, потому что такая работа по своей
    природе капает понемногу каждый день;
  * задачи по автоматизации весят вдвое больше обычных (AUTOMATION_WEIGHT):
    обычное functional testing — это 1.5–2 ч, а автоматизация тянет на 3–4 ч;
  * минимальный кусок — 0.5 ч; если свободных часов на всех не хватает, часть
    задач остаётся пустой (о чём вызывающий код предупредит).
"""

import re

QUARTER = 0.25
DAY_TARGET_Q = 32  # 8 ч в четвертях
STEP_Q = 2         # выдаём только кратное 0.5 ч
MIN_CHUNK_Q = 2    # 0.5 ч

# Задачи, которые логично вести понемногу каждый день, а не одним куском.
# Сравнение — по префиксу имени в нижнем регистре (в Excel они приходят с
# суффиксами вида "Bug verification P1-2, P2-4" / "Investigation issue 4").
SPREAD_PREFIXES = ("bug verification", "investigation issue")

# Задачи по автоматизации трудоёмче обычных — ищем по подстроке в имени
# ("Automation test maintenance", "Automation test update 2 test items", ...).
AUTOMATION_KEYWORDS = ("automation",)
AUTOMATION_WEIGHT = 2.0
DEFAULT_WEIGHT = 1.0


# Хвосты, которые правила именования навешивают на одну и ту же по сути работу:
#   "Investigation issue 2"            -> счётчик тикетов
#   "Bug verification P1-2, P2-4"      -> корзины приоритетов
#   "Automation test update 2 test items" -> число проверенных айтемов
# Всё это — одна строка табеля. А вот "Functional testing of story ID GT2-1662"
# заканчивается идентификатором, а не голым числом, и под шаблоны не попадает —
# такие задачи остаются разными, как и должны.
_SUFFIX_PATTERNS = (
    re.compile(r"\s+\d+(?:\.\d+)?\s+test\s+items$"),
    re.compile(r"\s+p\d+-\d+(?:\s*,\s*p\d+-\d+)*$"),
    re.compile(r"\s+\d+$"),
)


def normalize_task_name(name: str) -> str:
    """Ключ, по которому строки считаются одной и той же задачей."""
    n = " ".join((name or "").strip().lower().split())
    for pattern in _SUFFIX_PATTERNS:
        stripped = pattern.sub("", n)
        if stripped:  # не схлопываем имя в пустую строку
            n = stripped
    return n


def is_spread_task(name: str) -> bool:
    n = (name or "").strip().lower()
    return any(n.startswith(p) for p in SPREAD_PREFIXES)


def is_automation_task(name: str) -> bool:
    n = (name or "").strip().lower()
    return any(k in n for k in AUTOMATION_KEYWORDS)


def task_weight(name: str) -> float:
    """Доля задачи при делении свободных часов. Размазываемые задачи считаем
    обычными — им важна равномерность по дням, а не объём."""
    return AUTOMATION_WEIGHT if is_automation_task(name) else DEFAULT_WEIGHT


def hours_to_q(hours: float) -> int:
    """0.75 -> 3. Округляем к ближайшей четверти: значения из Excel приходят
    как float и могут быть вида 0.4999999999."""
    return int(round(float(hours) / QUARTER))


def q_to_hours(q: int) -> float:
    return round(q * QUARTER, 2)


def allocate(total_q: int, weights: list[float], step_q: int = STEP_Q,
             min_chunk_q: int = MIN_CHUNK_Q) -> list[int]:
    """Разделить total_q четвертей между задачами пропорционально весам.

    Возвращает список той же длины, сумма ровно total_q, каждое значение кратно
    step_q. Сначала всем выдаётся минимум (min_chunk_q), излишек делится по
    весам методом наибольшего остатка — так задача с весом 2 получает вдвое
    больше *сверх* минимума, а не вдвое больше «в среднем по больнице».

    Если на минимум всем не хватает — первые получают по min_chunk_q, остальные
    0. Неделимый хвост (total_q не кратно step_q) приклеивается к первой задаче:
    лучше одно нечётное число, чем разъехавшаяся сумма дня.
    """
    n = len(weights)
    if n == 0 or total_q <= 0:
        return [0] * n

    min_units = max(1, min_chunk_q // step_q)
    units, tail = divmod(total_q, step_q)

    if units < n * min_units:
        k = units // min_units
        alloc = [min_units * step_q] * k + [0] * (n - k)
        rest = total_q - k * min_units * step_q
        if rest:
            alloc[k - 1 if k else 0] += rest
        return alloc

    surplus = units - n * min_units
    total_w = sum(weights) or float(n)
    raw = [surplus * w / total_w for w in weights]
    extra = [int(x) for x in raw]
    left = surplus - sum(extra)
    # Наибольший остаток; при равных остатках вперёд идёт более тяжёлая задача.
    order = sorted(range(n), key=lambda i: (raw[i] - extra[i], weights[i]), reverse=True)
    for i in order[:left]:
        extra[i] += 1

    alloc = [(min_units + extra[i]) * step_q for i in range(n)]
    if tail:
        alloc[0] += tail
    return alloc


def spread(alloc_q: int, days: list[list[int]], step_q: int = STEP_Q) -> dict[int, int]:
    """Разложить alloc_q четвертей по дням кругами по step_q.

    `days` — [[колонка, свободно], ...]; меняется на месте (списываем ёмкость).
    Круговой обход даёт ровно то, что нужно для Bug verification: 2 ч уходят как
    0.5+0.5+0.5+0.5 по четырём дням, а не одним блоком в понедельник.
    """
    out: dict[int, int] = {}
    need = alloc_q
    i = 0
    n = len(days)
    while need > 0 and n and any(avail > 0 for _, avail in days):
        day = days[i % n]
        if day[1] > 0:
            take = min(step_q, need, day[1])
            out[day[0]] = out.get(day[0], 0) + take
            day[1] -= take
            need -= take
        i += 1
    return out


def spread_all(n_spread: int, days: list[list[int]], step_q: int = STEP_Q) -> list[dict[int, int]]:
    """Раздать ВСЮ оставшуюся ёмкость дней размазываемым задачам сразу.

    Ключевое отличие от последовательных вызовов `spread`: очередь задач общая и
    сквозная по дням, поэтому остаток, скопившийся в одном дне, делится между
    всеми, а не падает целиком на ту задачу, чья очередь подошла. Иначе третья
    по счёту задача получала единый кусок в 4.5 ч — ровно то, чего размазывание
    должно избегать.
    """
    outs: list[dict[int, int]] = [{} for _ in range(n_spread)]
    if n_spread <= 0:
        return outs
    turn = 0
    for day in days:
        col, avail = day[0], day[1]
        while avail > 0:
            take = min(step_q, avail)
            target = outs[turn % n_spread]
            target[col] = target.get(col, 0) + take
            avail -= take
            turn += 1
        day[1] = 0
    return outs


def split_counts(n_tasks: int, frees: list[int], min_chunk_q: int = MIN_CHUNK_Q) -> list[int]:
    """Сколько задач посадить в каждый день, чтобы куски вышли соразмерными.

    Сначала по одной задаче в каждый непустой день (иначе день останется
    недозаполненным), затем добавляем туда, где на задачу приходится больше
    всего часов. Больше, чем free/min_chunk, в день не сажаем — иначе кому-то
    достанется меньше 0.5 ч.
    """
    n_days = len(frees)
    counts = [0] * n_days
    if n_tasks <= 0 or n_days == 0:
        return counts

    caps = [f // min_chunk_q for f in frees]
    remaining = n_tasks
    for i, free in enumerate(frees):
        if free > 0 and caps[i] > 0 and remaining > 0:
            counts[i] = 1
            remaining -= 1

    while remaining > 0:
        best_i, best_ratio = None, -1.0
        for i, free in enumerate(frees):
            if free <= 0 or counts[i] >= caps[i]:
                continue
            ratio = free / (counts[i] + 1)
            if ratio > best_ratio:
                best_i, best_ratio = i, ratio
        if best_i is None:
            break
        counts[best_i] += 1
        remaining -= 1
    return counts


def distribute(total_q: int, caps: list[int], step_q: int = STEP_Q) -> list[int]:
    """Разложить total_q четвертей по «корзинам» пропорционально их ёмкости.

    Каждая доля кратна step_q и не превышает свою ёмкость; сумма — ровно total_q
    (или сумма ёмкостей, если запрошено больше). Нужно, чтобы поделить бюджет
    блочных задач между днями, не залезая в часы, зарезервированные под
    размазываемые задачи.
    """
    n = len(caps)
    out = [0] * n
    if n == 0 or total_q <= 0:
        return out

    total_cap = sum(caps)
    if total_cap <= 0:
        return out
    total_q = min(total_q, total_cap)

    units_total, tail = divmod(total_q, step_q)
    raw = [units_total * c / total_cap for c in caps]
    base = [min(int(x), caps[i] // step_q) for i, x in enumerate(raw)]
    left = units_total - sum(base)
    order = sorted(range(n), key=lambda i: raw[i] - base[i], reverse=True)
    while left > 0:
        progressed = False
        for i in order:
            if left <= 0:
                break
            if (base[i] + 1) * step_q <= caps[i]:
                base[i] += 1
                left -= 1
                progressed = True
        if not progressed:
            break

    out = [b * step_q for b in base]
    if tail:
        for i in range(n):
            if out[i] + tail <= caps[i]:
                out[i] += tail
                break
    return out


def plan_week(used_by_col: dict[int, float], active_cols: list[int], task_names: list[str],
              day_target_q: int = DAY_TARGET_Q,
              spread_all_tasks: bool = False) -> tuple[list[dict[int, float]], int]:
    """Спланировать добивку одной недели.

    used_by_col — сколько часов уже стоит в каждой колонке (все строки блока);
    active_cols — колонки дней, которые можно заполнять (обрезка по месяцу
                  выкидывает дни соседнего месяца);
    task_names  — имена строк, доступных под раскладку, в порядке их следования
                  в файле. Имя решает всё: размазывать задачу по неделе или
                  класть одним куском, и какой у неё вес.

    Возвращает (план, свободно_четвертей_всего), где план — список по задачам:
    {колонка: часы}. Задачи, которым ничего не досталось, дают пустой dict.
    """
    n_tasks = len(task_names)
    free_by_day = []
    for col in active_cols:
        free = day_target_q - hours_to_q(used_by_col.get(col, 0) or 0)
        free_by_day.append((col, max(free, 0)))
    total_free = sum(free for _, free in free_by_day)

    if total_free <= 0 or n_tasks <= 0:
        return [{} for _ in range(n_tasks)], total_free

    weights = [task_weight(name) for name in task_names]
    packed: list[dict[int, int]] = [{} for _ in range(n_tasks)]
    days = [[col, free] for col, free in free_by_day]

    # Будущая неделя — это план, а не отчёт: там нет «в среду я делал вот это»,
    # поэтому каждая задача честно размазывается по всем дням. Правило «одна
    # задача — один день» защищает достоверность уже отработанных недель, а
    # здесь оно только рисует неправдоподобные глыбы по 6.5 ч.
    if spread_all_tasks:
        for col, free in free_by_day:
            if free <= 0:
                continue
            for i, q in enumerate(allocate(free, weights)):
                if q > 0:
                    packed[i][col] = packed[i].get(col, 0) + q
        return [{col: q_to_hours(q) for col, q in row.items() if q > 0} for row in packed], total_free

    spread_idx = [i for i, name in enumerate(task_names) if is_spread_task(name)]
    block_idx = [i for i in range(n_tasks) if i not in set(spread_idx)]

    # Доля блочных задач в общем котле — по весам. Остальное резервируем под
    # размазываемые: они пойдут последними и подберут всё, что осталось в днях.
    share = allocate(total_free, weights)
    block_budget = sum(share[i] for i in block_idx)

    # 1. Блочные задачи: каждая целиком в один день, без переноса на следующий.
    frees = [avail for _, avail in days]
    counts = split_counts(len(block_idx), frees, MIN_CHUNK_Q)
    # Бюджет делим только между днями, куда реально сели задачи — иначе часы
    # зависнут в дне, которому задач не досталось.
    host_days = [d for d in range(len(days)) if counts[d] > 0]
    day_budgets = distribute(block_budget, [frees[d] for d in host_days])
    cursor = 0
    for pos, d in enumerate(host_days):
        col = days[d][0]
        budget = day_budgets[pos]
        day_tasks = block_idx[cursor:cursor + counts[d]]
        cursor += counts[d]
        if not day_tasks or budget <= 0:
            continue
        for i, q in zip(day_tasks, allocate(budget, [weights[i] for i in day_tasks])):
            if q > 0:
                packed[i] = {col: q}
                days[d][1] -= q

    # 2. Размазываемые задачи добирают всё, что осталось, кругами по дням. Они
    #    же закрывают дни, которым не хватило блочных задач — поэтому идут
    #    последними, а не первыми.
    if spread_idx and sum(avail for _, avail in days) > 0:
        for i, row in zip(spread_idx, spread_all(len(spread_idx), days)):
            packed[i] = row

    return [{col: q_to_hours(q) for col, q in row.items() if q > 0} for row in packed], total_free
