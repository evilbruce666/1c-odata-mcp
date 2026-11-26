"""Выдуманные данные и вопросы замера. Одни и те же ответы получают все участники сравнения; настоящей 1С и сети нет."""
import json, os, re

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def load_tools():
    """18 инструментов набора для Бухгалтерии 3.0 из docs/CONNECTING.md: описания и схемы — из manifest.json."""
    doc = open(os.path.join(ROOT, "docs/CONNECTING.md"), encoding="utf-8").read()
    names = re.search(r"TOOLS_ALLOW=([\w.,]+)", doc[doc.index("Набор для Бухгалтерии 3.0:"):]).group(1).split(",")
    manifest = {t["name"]: t for t in json.load(open(os.path.join(ROOT, "manifest.json"), encoding="utf-8"))["tools"]}
    return [{"name": n, "description": manifest[n]["description"], "inputSchema": manifest[n]["inputSchema"]} for n in names]

FIRMS=["Ромашка","Василёк","Лютик","Клевер","Астра","Пион","Ирис","Мак","Георгин","Нарцисс","Флокс","Левкой"]
def rows(n,f): return [f(i) for i in range(n)]
def result(name,args):
    if name=="read.system.list_databases": return {"databases":[{"name":"demo","default":True,"configuration":"Бухгалтерия 3.0"}]}
    if name=="read.system.list_organizations": return {"organizations":[{"name":"ООО «Тестовая»","inn":"0000000000"}]}
    if name=="read.system.health_check": return {"ok":True,"database":"demo"}
    if name=="read.analytics.get_cash_balance":
        return {"asOf":"2026-11-26","accounts":[{"account":"51","name":"Расчётный счёт, Банк А","balance":1612400},{"account":"51","name":"Расчётный счёт, Банк Б","balance":187650},{"account":"50","name":"Касса","balance":45250}],"totals":{"available":1845300,"special":0,"inTransit":0},"currency":"RUB"}
    if name=="read.analytics.get_debtors":
        r=[{"counterparty":"ООО «Ромашка»","debt":940000,"overdue":310000,"notDue":630000,"oldestDays":47}]
        r+=rows(11,lambda i:{"counterparty":f"ООО «{FIRMS[i+1]}»","debt":420000-i*31000,"overdue":(i%3)*12500,"notDue":420000-i*31000-(i%3)*12500,"oldestDays":12+i*3})
        return {"asOf":"2026-11-26","total":sum(x["debt"] for x in r),"overdue":sum(x["overdue"] for x in r),"rows":r,"truncated":False}
    if name in ("read.analytics.get_sales_breakdown","read.analytics.get_sales_report"):
        top=[("Ромашка",1250000),("Василёк",870500),("Лютик",612000)]+[(FIRMS[i],480000-i*29000) for i in range(3,12)]
        return {"from":"2026-10-01","to":"2026-10-31","groupBy":"counterparty","total":sum(a for _,a in top),"rows":[{"counterparty":f"ООО «{n}»","revenue":a,"documents":max(1,a//90000)} for n,a in top],"truncated":False}
    if name=="read.analytics.get_unpaid_invoices":
        r=[{"number":"ТД-000214","date":"2026-10-28","counterparty":"ООО «Василёк»","amount":215000,"paid":0,"rest":215000},{"number":"ТД-000231","date":"2026-11-09","counterparty":"ООО «Клевер»","amount":98400,"paid":40000,"rest":58400}]
        return {"asOf":"2026-11-26","totalRest":273400,"rows":r,"truncated":False}
    if name=="read.accounting.check_month_close":
        return {"month":"2026-09","closed":False,"checks":[{"check":"Закрытие месяца","status":"error","message":"Регламентная операция «Закрытие месяца» не выполнена"},{"check":"Непроведённые документы","status":"warning","count":3},{"check":"Красное сальдо","status":"warning","message":"41.01: отрицательный остаток по 1 позиции"},{"check":"Предыдущие месяцы","status":"ok"},{"check":"Остатки на 90/91","status":"ok"}]}
    return {"rows":[],"note":"В тестовой базе по этому запросу данных нет."}
def text(name,args): return json.dumps(result(name,args),ensure_ascii=False)
QUESTIONS=[  # вопрос, что обязано быть в ответе
 ("Сколько у нас сейчас денег на счетах и в кассе?",["1845300"]),
 ("Кто нам должен больше всех и какая часть этого долга просрочена?",["Ромашка","940000","310000"]),
 ("Назови трёх крупнейших покупателей по выручке за октябрь 2026 года.",["Ромашка","Василёк","Лютик"]),
 ("Проверь, закрыт ли сентябрь 2026 года, и что с ним не так.",["не","3","41"]),
 ("Дай короткую сводку: сколько денег, кто должен больше всех и какие счета покупателей не оплачены.",["1845300","Ромашка","215000"]),
]
import re
def has(answer,facts):
    t=re.sub(r"(?<=\d)[\s  ,.'](?=\d{3}(\D|$))","",answer)
    return all(f in t for f in facts)
SYSTEM="Ты помощник по 1С. Отвечай по-русски, кратко, только по данным инструментов. Сегодня 2026-11-26."
