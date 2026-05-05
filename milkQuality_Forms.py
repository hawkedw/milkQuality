# milkQuality_Forms.py
# Импорт и выгрузка форм 1 / 2 / 5 ArcGIS в Excel

import sys, os, json, datetime
import requests
import win32com.client as win32

LOG_PATH = os.path.join(os.path.dirname(__file__), "milkQuality_Forms.log")

def log(msg: str):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


DIRTY_ALIAS = "Dirty"
SHEET_F1 = "Форма 1"
SHEET_F2 = "Форма 2"
SHEET_F5 = "Форма 5"

URL_F1 = "https://maps.ekoniva-apk.org/arcgis/rest/services/quality_surveys/milkQuality/FeatureServer/2"
URL_F2 = "https://maps.ekoniva-apk.org/arcgis/rest/services/quality_surveys/milkQuality/FeatureServer/3"
URL_F5 = "https://maps.ekoniva-apk.org/arcgis/rest/services/quality_surveys/milkQuality/FeatureServer/9"

SORT_F1_FIELD = "timeStart"
SORT_F2_FIELD = "milktruk_departure_time"

# ---------- FIELD MAPS ----------
FIELDS_F1 = [
    {"n":"timeStart","alias":"Время начала доения","type":"DATE"},
    {"n":"timeEnd","alias":"Время окончания доения","type":"DATE"},
    {"n":"cattle_complex_tank","alias":"ЖК","type":"TEXT"},
    {"n":"reservoir_num","alias":"Номер танка","type":"INT"},
    {"n":"milking_num","alias":"Доение","type":"TEXT"},
    {"n":"milking_remaining_yesterday","alias":"Проверка остатков с предыдущего дня","type":"TEXT"},
    {"n":"milk_weight","alias":"Вес, кг","type":"NUMBER"},
    {"n":"reservoir_num_remain_yesterd","alias":"Номера танков (с предыдущего дня)","type":"TEXT"},
    {"n":"milking_remaining","alias":"Остаток на начало доения","type":"TEXT"},
    {"n":"milk_weight_remaining","alias":"Вес, кг (остаток)","type":"NUMBER"},
    {"n":"reservoir_num_remaining","alias":"Номера танков (с остатком)","type":"TEXT"},
    {"n":"timeEnd_remaining","alias":"Время окончания доения (остаток)","type":"DATE"},
    {"n":"milk_fat_2","alias":"Массовая доля жира (2е измерение), %","type":"NUMBER"},
    {"n":"milk_fat_3","alias":"Массовая доля жира (3е измерение), %","type":"NUMBER"},
    {"n":"milk_protein_2","alias":"Массовая доля белка (2е измерение), %","type":"NUMBER"},
    {"n":"milk_protein_3","alias":"Массовая доля белка (2е измерение), %","type":"NUMBER"},
    {"n":"milk_acidity","alias":"Кислотность, °Т","type":"NUMBER"},
    {"n":"milk_density","alias":"Плотность, кг/м3","type":"INT"},
    {"n":"milk_termstab","alias":"Термоустойчивость","type":"INT"},
    {"n":"milk_purity","alias":"Группа чистоты","type":"INT"},
    {"n":"milk_temp_sensor","alias":"Температура на датчике, °С","type":"NUMBER"},
    {"n":"milk_temp_therm","alias":"Температура на термометре, °С","type":"NUMBER"},
    {"n":"milk_fat_inTank","alias":"Жир/кг в каждом танке","type":"NUMBER"},
    {"n":"milk_protein_inTank","alias":"Бел/кг в каждом танке","type":"NUMBER"},
    {"n":"personal_name_tank","alias":"ФИО заполнившего акт проверки","type":"TEXT"},
    {"n":"year_tank","alias":"год","type":"INT"},
    {"n":"month_tank","alias":"месяц","type":"INT"},
    {"n":"day_tank","alias":"день","type":"INT"},
    {"n":"joinIndex_tank","alias":"индекс","type":"TEXT"},
    {"n":"created_user","alias":"created_user","type":"TEXT"},
    {"n":"created_date","alias":"created_date","type":"DATE"},
    {"n":"last_edited_user","alias":"last_edited_user","type":"TEXT"},
    {"n":"last_edited_date","alias":"last_edited_date","type":"DATE"},
    {"n":"OBJECTID","alias":"OBJECTID","type":"OID"},
    {"n":"parentglobalid","alias":"parentglobalid","type":"GUID"},
    {"n":"GlobalID","alias":"GlobalID","type":"GUID"}
]

FIELDS_F2 = [
    {"n":"milktruk_departure_time","alias":"Время отправления молоковоза","type":"DATE"},
    {"n":"cattle_complex_milktruk","alias":"ЖК","type":"TEXT"},
    {"n":"counteragent_milktruk","alias":"Контрагент","type":"TEXT"},
    {"n":"invoice_parent","alias":"№ ТТН","type":"TEXT"},
    {"n":"milktruk_section_num","alias":"Номер секции","type":"INT"},
    {"n":"milk_fat_invoice","alias":"М.д.жира (Результат в ТТН)","type":"NUMBER"},
    {"n":"milk_fat_milkfactory","alias":"М.д.жира с завода","type":"NUMBER"},
    {"n":"milk_fat_difference","alias":"М.д.жира, расхождение","type":"NUMBER"},
    {"n":"milk_protein_invoice","alias":"М.д.белка (Результат в ТТН)","type":"NUMBER"},
    {"n":"milk_protein_milkfactory","alias":"М.д.белка с завода","type":"NUMBER"},
    {"n":"milk_protein_difference","alias":"М.д.белка, расхождение","type":"NUMBER"},
    {"n":"milk_weight_ekoniva","alias":"Вес (отгружено)","type":"NUMBER"},
    {"n":"milk_weight_milkfactory","alias":"Вес с завода, кг","type":"NUMBER"},
    {"n":"milk_weight_difference","alias":"Вес, расхождение, кг","type":"NUMBER"},
    {"n":"reservoir_num_1","alias":"Номер танка 1","type":"INT"},
    {"n":"reservoir_num_2","alias":"Номер танка 2","type":"INT"},
    {"n":"milk_fat_2","alias":"М.д.жира (2е изм), %","type":"NUMBER"},
    {"n":"milk_fat_3","alias":"М.д.жира (3е изм), %","type":"NUMBER"},
    {"n":"milk_protein_2","alias":"Массовая доля белка (2е изм), %","type":"NUMBER"},
    {"n":"milk_protein_3","alias":"Массовая доля белка (2е изм), %","type":"NUMBER"},
    {"n":"personal_name_milktruk","alias":"ФИО заполнившего акт","type":"TEXT"},
    {"n":"year_milktruk","alias":"год","type":"INT"},
    {"n":"month_milktruk","alias":"месц","type":"INT"},
    {"n":"day_milktruk","alias":"день","type":"INT"},
    {"n":"joinIndex_milktruk","alias":"индекс","type":"TEXT"},
    {"n":"created_user","alias":"created_user","type":"TEXT"},
    {"n":"created_date","alias":"created_date","type":"DATE"},
    {"n":"last_edited_user","alias":"last_edited_user","type":"TEXT"},
    {"n":"last_edited_date","alias":"last_edited_date","type":"DATE"},
    {"n":"OBJECTID","alias":"OBJECTID","type":"OID"},
    {"n":"GlobalID","alias":"GlobalID","type":"GUID"},
    {"n":"parentglobalid","alias":"parentglobalid","type":"GUID"}
]

# --- Форма 5 (колонки) ---
# ВАЖНО: для виртуальных колонок итогов добавляем id, чтобы строить шапку по ключам.

FIELDS_F5 = [
    {"n": "personal_name", "alias": "ФИО заполнившего чек-лист", "type": "TEXT"},
    {"n": "cattle_complex_parent", "alias": "ЖК", "type": "TEXT"},
    {"n": "dateForm5", "alias": "Дата", "type": "DATE"},
    {"n": "counteragents", "alias": "Контрагент", "type": "TEXT"},

    # block_1: Контроль молока в танках
    {"n": "method", "alias": "Методика определения", "type": "TEXT"},
    {"n": "method_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "brand", "alias": "Марка, условия хранения тест-полосок", "type": "TEXT"},
    {"n": "brand_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "logBook", "alias": "Журнал учета результатов измерения", "type": "TEXT"},
    {"n": "logBook_s", "alias": "Оценка, балл", "type": "INT"},

    {"n": "dishes", "alias": "Наличие посуды и оборудования (цилиндр, ареометр),", "type": "TEXT"},
    {"n": "dishes_s", "alias": "Оценка, балл", "type": "INT"},

    {"n": "logBook_logs", "alias": "Наличие записей в Журнале контроля отгрузки", "type": "TEXT"},
    {"n": "logBook_logs_s", "alias": "Оценка, балл", "type": "INT"},

    {"n": "reagentsAcid", "alias": "Наличие реактивов, посуды", "type": "TEXT"},
    {"n": "reagentsAcid_s", "alias": "Оценка, балл", "type": "INT"},

    {"n": "reagentsAlco", "alias": "Наличие реактивов, посуды", "type": "TEXT"},
    {"n": "reagentsAlco_s", "alias": "Оценка, балл", "type": "INT"},

    {"n": "equipment", "alias": "Наличие оборудования", "type": "TEXT"},
    {"n": "equipment_s", "alias": "Оценка, балл", "type": "INT"},

    {"n": "thermometers", "alias": "Наличие термометров", "type": "TEXT"},
    {"n": "thermometers_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "temperature", "alias": "Контроль температуры молока в танке, периодичность", "type": "TEXT"},
    {"n": "temperature_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "thermometersMaintenance", "alias": "ТО термодатчиков, периодичность, ответственный", "type": "TEXT"},
    {"n": "thermometersMaintenance_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "temperatureLogbook", "alias": "Журнал фиксации данных температуры молока в танке (фиксация в журнале учета молока)", "type": "TEXT"},
    {"n": "temperatureLogbook_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "shipmentTemperature", "alias": "Соблюдение температурного режима отгрузки, температура отгрузки", "type": "TEXT"},
    {"n": "shipmentTemperature_s", "alias": "Оценка, балл", "type": "INT"},
    {"id": "total_block1", "n": None, "alias": "Итого", "type": "INT"},   # AE

    # block_2: Подготовка к отгрузке молока
    {"n": "sanitaryCondition", "alias": "Санитарное состояние молочного оборудования (резинки, насосы,уплотнители)", "type": "TEXT"},
    {"n": "sanitaryCondition_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "storingRules", "alias": "Соблюдение правил хранения отгрузочного шланга и подготовка к отгрузке", "type": "TEXT"},
    {"n": "storingRules_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "milkStirring", "alias": "Включение перемешивания молока перед отгрузкой на 30 мин", "type": "TEXT"},
    {"n": "milkStirring_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "tankerSealsInspection", "alias": "Осмотр опломбировки автоцистерны (наличие пломб, № пломб, сравнение с помывочным листом)", "type": "TEXT"},
    {"n": "tankerSealsInspection_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "cleanlinessCheck", "alias": "Проверка чистоты автоцистерны молоковоза", "type": "TEXT"},
    {"n": "cleanlinessCheck_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "closedDrainValve", "alias": "Наличие закрытого крана слива", "type": "TEXT"},
    {"n": "closedDrainValve_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "washingSheet", "alias": "Предоставление помывочного листа (№, дата сан. обработки)", "type": "TEXT"},
    {"n": "washingSheet_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "waybill", "alias": "Путевой лист водителя (№, дата) наличие", "type": "TEXT"},
    {"n": "waybill_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "calibration", "alias": "Тарировка (№, дата свидетельства), наличие", "type": "TEXT"},
    {"n": "calibration_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "powerOfAttorney", "alias": "Доверенность, наличие", "type": "TEXT"},
    {"n": "powerOfAttorney_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "nonConformityReport", "alias": "Составление акта несоответствия в случае несоответствия по состоянию АМЦ или сопроводительных документов", "type": "TEXT"},
    {"n": "nonConformityReport_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "invoiceCopy", "alias": "Наличие копии ТТН, помывочного листа, срок хранения", "type": "TEXT"},
    {"n": "invoiceCopy_s", "alias": "Оценка, балл", "type": "INT"},
    {"id": "total_block2", "n": None, "alias": "Итого", "type": "INT"},   # BD

    # block_3: Отгрузка и отбор проб
    {"n": "loadingControl", "alias": "Контроль полноты загрузки секций по тарировке", "type": "TEXT"},
    {"n": "loadingControl_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "sealIntegrityConfirmed", "alias": "Опломбировку секций и крана слива осуществляет слесарь (наличие пломб, № пломб в ттн, отсутствие возможности доступа без повреждения пломб)", "type": "TEXT"},
    {"n": "sealIntegrityConfirmed_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "specializedEquipment", "alias": "Наличие специализированного оборудования и посуды для отбора (указать, чем производят отбор, как отбирают)", "type": "TEXT"},
    {"n": "specializedEquipment_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "samplingEquipment", "alias": "Санитарная обработка пробоотборника (черпака)", "type": "TEXT"},
    {"n": "samplingEquipment_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "sampledBy", "alias": "Отбор проб выполняет", "type": "TEXT"},
    {"n": "sampledBy_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "sampleContainerVolume", "alias": "Емкость для хранения и отбора проб (объем, материал)", "type": "TEXT"},
    {"n": "sampleContainerVolume_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "controlSampleStorage", "alias": "Хранение и маркировка контрольной пробы (условия, продолжительность хранения, полнота маркировки)", "type": "TEXT"},
    {"n": "controlSampleStorage_s", "alias": "Оценка, балл", "type": "INT"},
    {"id": "total_block3", "n": None, "alias": "Итого", "type": "INT"},   # BS

    # block_4: Проведение измерений ФХП молока-сырья
    {"n": "analyzerName", "alias": "Наименование анализатора, заводской номер", "type": "TEXT"},
    {"n": "analyzerName_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "printer", "alias": "Принтер для печати чеков (наличие)", "type": "TEXT"},
    {"n": "printer_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "washRegimeCompliant", "alias": "Правильное обслуживание прибора (соблюдение режимов мойки)", "type": "TEXT"},
    {"n": "washRegimeCompliant_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "washInstructionVersion", "alias": "Инструкция по мойке прибора (наличие, номер версии, дата)", "type": "TEXT"},
    {"n": "washInstructionVersion_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "measurementProcedureNotes", "alias": "Правильность выполнения измерений (перемешивание пробы, отсутствие пузырьков, измерение показателей 3 раза", "type": "TEXT"},
    {"n": "measurementProcedureNotes_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "milkShipmentControlLog", "alias": "Журнал фиксации данных (журнал контроля отгрузки молока)", "type": "TEXT"},
    {"n": "milkShipmentControlLog_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "replicateDiffWithin005", "alias": "Разница между 2 и 3 измерением, а также между танком и секциями не превышает 0,05%,", "type": "TEXT"},
    {"n": "replicateDiffWithin005_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "thermoReceiptStorage", "alias": "Хранение термочеков, в случае отсутствия термопринтера, данные прибора фиксируют за подписью", "type": "TEXT"},
    {"n": "thermoReceiptStorage_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "analysisResponsible", "alias": "Ответственый за анализ", "type": "TEXT"},
    {"n": "analysisResponsible_s", "alias": "Оценка, балл", "type": "INT"},

    {"n": "analyzerFat", "alias": "Анализатор", "type": "TEXT"},
    {"n": "analyzerFat_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "laboratoryFat", "alias": "Лаборатория", "type": "TEXT"},
    {"n": "laboratoryFat_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "differenceFat", "alias": "Разница", "type": "TEXT"},
    {"n": "differenceFat_s", "alias": "Оценка, балл", "type": "INT"},

    {"n": "analyzerProtein", "alias": "Анализатор", "type": "TEXT"},
    {"n": "analyzerProtein_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "laboratoryProtein", "alias": "Лаборатория", "type": "TEXT"},
    {"n": "laboratoryProtein_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "differenceProtein", "alias": "Разница", "type": "TEXT"},
    {"n": "differenceProtein_s", "alias": "Оценка, балл", "type": "INT"},
    {"id": "total_block4", "n": None, "alias": "Итого", "type": "INT"},   # CX

    # block_5: Количественный учет отгружаемого молока
    {"n": "primaryMilkMeter", "alias": "Основной прибор количественного учета молока, по которому проводятся отгрузки", "type": "TEXT"},
    {"n": "additionalMeters", "alias": "Наличие дополнительных приборов учета на хозяйстве (указать какие)", "type": "TEXT"},
    {"n": "invoiceQuantityMethod", "alias": "Количество по счетчику/линейке/весам (указать способ учета, по которому заполняли ТТН по данной отгрузке)", "type": "TEXT"},
    {"n": "additionalMeterQuantity", "alias": "Количество по дополнительному прибору учета (сверка) при наличии", "type": "TEXT"},
    {"n": "meterDiffKg", "alias": "Разница показаний приборов учета (путем вычитания из показаний основного прибора показаний дополнительного) в кг и % от общей отгруженной массы", "type": "TEXT"},
    {"n": "meterDiffKg_s", "alias": "Оценка, балл", "type": "INT"},
    {"n": "conversionDensity", "alias": "Плотность, применяемая для пересчета", "type": "INT"},
    {"n": "conversionDensity_s", "alias": "Оценка, балл", "type": "INT"},
    {"id": "total_block5", "n": None, "alias": "Итого", "type": "INT"},   # DG

    {"id": "total_all", "n": None, "alias": "Итого, всего", "type": "INT"},  # DH

    # смещения под служебные поля (как у тебя было)
    {"n": None, "alias": "", "type": "INT"},
    {"n": None, "alias": "", "type": "INT"},
    {"n": None, "alias": "", "type": "INT"},

    {"n": "created_user", "alias": "created_user", "type": "TEXT"},
    {"n": "created_date", "alias": "created_date", "type": "DATE"},
    {"n": "last_edited_user", "alias": "last_edited_user", "type": "TEXT"},
    {"n": "last_edited_date", "alias": "last_edited_date", "type": "DATE"},
    {"n": "OBJECTID", "alias": "OBJECTID", "type": "OID"},
]

# --- Форма 5: схема групп/подгрупп для строк 1-2 (ключи = f["n"]) ---

FORM5_GROUPS = [
    {"text": "Общая информация", "start": "personal_name", "end": "counteragents"},
    {"text": "Контроль молока в танках", "start": "method", "end": "shipmentTemperature_s", "end_col": 31},      # AE
    {"text": "Подготовка к отгрузке молока", "start": "sanitaryCondition", "end": "invoiceCopy_s", "end_col": 56},# BD
    {"text": "Отгрузка и отбор проб", "start": "loadingControl", "end": "controlSampleStorage_s", "end_col": 71}, # BS
    {"text": "Проведение измерений ФХП молока-сырья", "start": "analyzerName", "end": "differenceProtein_s", "end_col": 102}, # CX
    {"text": "Количественный учет отгружаемого молока", "start": "primaryMilkMeter", "end": "conversionDensity_s", "end_col": 111}, # DG
]

FORM5_SUBGROUPS = [
    # block_1
    {"text": "Определение антибиотиков", "start": "method", "end": "logBook_s"},
    {"text": "Плотность при фактической температуре согласно ТП29", "start": "dishes", "end": "dishes_s"},
    {"text": "Массовая доля жира и белка на анализаторе (Р1 1)", "start": "logBook_logs", "end": "logBook_logs_s"},
    {"text": "Кислотность согласно ТП3", "start": "reagentsAcid", "end": "reagentsAcid_s"},
    {"text": "Термоустойчивость по алкогольной пробе по ТП4", "start": "reagentsAlco", "end": "reagentsAlco_s"},
    {"text": "Группа чистоты по ТП8", "start": "equipment", "end": "equipment_s"},
    {"text": "Температура", "start": "thermometers", "end": "shipmentTemperature_s"},

    # block_2
    {"text": "Санитарное состояние (отгрузочный шланг, насос)", "start": "sanitaryCondition", "end": "storingRules_s"},
    {"text": "Подготовка молока", "start": "milkStirring", "end": "milkStirring_s"},
    {"text": "Состояние автоцистерны", "start": "tankerSealsInspection", "end": "closedDrainValve_s"},
    {"text": "Сопроводительные документы автоцистерны", "start": "washingSheet", "end": "nonConformityReport_s"},
    {"text": "Хранение сопроводительных документов", "start": "invoiceCopy", "end": "invoiceCopy_s"},

    # block_3
    {"text": "Полнота загрузки секций", "start": "loadingControl", "end": "loadingControl_s"},
    {"text": "Опломбировка автоцистерны", "start": "sealIntegrityConfirmed", "end": "sealIntegrityConfirmed_s"},
    {"text": "Отбор проб", "start": "specializedEquipment", "end": "sampledBy_s"},
    {"text": "Хранение проб", "start": "sampleContainerVolume", "end": "controlSampleStorage_s"},

    # block_4
    {"text": "Работа анализатора", "start": "analyzerName", "end": "analysisResponsible_s"},
    {"text": "Массовая доля жира%", "start": "analyzerFat", "end": "differenceFat_s"},
    {"text": "Массовая доля белка%", "start": "analyzerProtein", "end": "differenceProtein_s"},

    # block_5
    {"text": "Количество отгруженного молока", "start": "primaryMilkMeter", "end": "conversionDensity_s"},
]


# ---------- ARC / AUTH ----------
TOKEN_URL = "https://maps.ekoniva-apk.org/portal/sharing/rest/generateToken"


ARC_USERNAME_ENV = "ARCGIS_QUALITY_USER"
ARC_PASSWORD_ENV = "ARCGIS_QUALITY_PASS"


def get_token() -> str:
    """Получить токен ArcGIS Portal.

    Логика:
    - имя пользователя и пароль берутся из переменных окружения
      ARCGIS_QUALITY_USER и ARCGIS_QUALITY_PASS;
    - никакие учётные данные не хардкодятся в репозитории/файлах;
    - client=referer, referer указываем на портал maps.ekoniva-apk.org.
    """
    username = os.environ.get(ARC_USERNAME_ENV)
    password = os.environ.get(ARC_PASSWORD_ENV)
    if not username or not password:
        raise RuntimeError(
            "Не заданы переменные окружения "
            f"{ARC_USERNAME_ENV}/{ARC_PASSWORD_ENV} с учётными данными."
        )

    payload = {
        "username": username,
        "password": password,
        "client": "referer",
        "referer": "https://maps.ekoniva-apk.org",
        "expiration": 60,
        "f": "json",
    }
    resp = requests.post(TOKEN_URL, data=payload, timeout=30)
    js = resp.json()
    tok = js.get("token")
    if not tok:
        # Логируем ответ для отладки, если что-то пойдет не так
        log(f"Token error: {js}")
        raise RuntimeError(f"Token error: {js}")
    return tok


# ---------- HELPERS ----------

EPOCH = datetime.datetime(1970, 1, 1)

# Локальный оффсет (MSK). Используется ТОЛЬКО для отображения/ввода локального времени.
OFFSET = datetime.timedelta(hours=3)

# Excel / OLE Automation epoch (важно: именно 1899-12-30)
EXCEL_EPOCH = datetime.datetime(1899, 12, 30)

# ----------------- DATE FIXES (F1/F2) -----------------

# Проблемный алиас в Ф1 может отличаться пробелом перед скобкой, поэтому сравниваем "нормализовано".
_FORCE_DATE