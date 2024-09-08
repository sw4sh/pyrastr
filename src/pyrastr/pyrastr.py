from win32com.client import Dispatch, WithEvents
from .dynamic import RastrDynamic
from typing import List, Literal
import warnings
import logging
import os

logging.basicConfig(format=u"[%(asctime)s] [%(name)s] [%(levelname)s] - %(message)s", level=logging.NOTSET)


class RastrEvents:
    """
    Класс обработчика событий Rastr. Набор используемых обработчиков определяется при инциализации объекта.
    С помощью наследования и переопределения методов данного класса можно построить собственную
    логику обработки событий.
    """

    def __init__(self):
        """
        Класс обработчика событий Rastr. Набор используемых обработчиков определяется при инциализации объекта.
        С помощью наследования и переопределения методов данного класса можно построить собственную
        логику обработки событий.
        """
        pass

    @staticmethod
    def OnChangeData(hint, tabl, col, row) -> None:
        __RastrHintInfo = {
            0: 'HINTR_CHANGE_ALL',
            1: 'HINTR_CHANGE_COL',
            2: 'HINTR_CHANGE_ROW',
            3: 'HINTR_CHANGE_DATA',
            4: 'HINTR_ADD_ROW',
            5: 'HINTR_DELETE_ROW',
            6: 'HINTR_INS_ROW',
            7: 'HINTR_CHANGE_TABL'
        }
        logging.info(f"({__RastrHintInfo.get(hint)}) Изменены данные таблицы [{tabl}] в столбце [{col}], "
                     f"row_id [{row}].")

    @staticmethod
    def Onprot(message) -> None:
        logging.info(f"Onprot: {message}")

    @staticmethod
    def OnCommandMain(comm, p1, p2, pp, p_val):
        logging.info(f"OnCommandMain: {comm, p1, p2, pp, p_val}")

    @staticmethod
    def OnUndo(type_, level):
        logging.info(f"OnUndo: {type_, level}")

    @staticmethod
    def OnHistoryChange(type_):
        logging.info(f"OnHistoryChanged: {type_}")

    @staticmethod
    def OnLog(code, level, id_, name, index, description, form_name) -> None:
        __RastrCodeInfo = {
            0: 'System Error',
            1: 'Failed',
            2: 'Error',
            3: 'Warning',
            4: 'Message',
            5: 'Info',
            6: 'Stage Open',
            7: 'Stage Close',
            8: 'Enter Default',
            9: 'Reset',
            10: 'None'
        }

        __loggingCodes = {
            'System Error': logging.CRITICAL,
            'Failed': logging.CRITICAL,
            'Error': logging.ERROR,
            'Warning': logging.WARNING,
            'Message': logging.INFO,
            'Info': logging.INFO,
            'Stage Open': logging.DEBUG,
            'Stage Close': logging.DEBUG,
            'Enter Default': logging.DEBUG,
            'Reset': logging.DEBUG,
            'None': logging.NOTSET
        }
        logging.log(
            level=__loggingCodes.get(__RastrCodeInfo.get(code)),
            msg=description)


class RastrMacroStudio:
    """
    Класс для взаимодействия с макростудией Rastrwin3
    """

    def __init__(self, astra):
        """
        Класс для взаимодействия с макростудией Rastrwin3
        """
        self._Astra = astra

    def run(self, path: str = None, script: str = None, parameters: str = ""):
        if path is not None:
            return self._Astra.ExecMacroPath(path, parameters)
        elif script is not None:
            return self._Astra.ExecMacroSource(script, parameters)
        else:
            raise FileNotFoundError(
                f"Для выполнения скрипта в макростудии Rastrwin3 должен быть задан путь или текст скрипта.")


class Rastr:
    """
    Базовый класс для взаимодействия с RastrWin3
    """

    def __init__(self, with_events: bool = True, event_handler=RastrEvents):
        """
        Базовый класс для взаимодействия с RastrWin3
        """
        self._Astra = Dispatch('Astra.Rastr')
        self.Tables = RastrTables(self._Astra.Tables)
        self.Dynamic = RastrDynamic(self._Astra.FWDynamic())
        self.MacroStudio = RastrMacroStudio(self._Astra)
        self._pars = ["", "p", "z", "c", "r", "i"]
        self._calc_codes = {0: "AST_OK", 1: "AST_NB", 2: "AST_REPT"}
        if with_events:
            WithEvents(self._Astra, event_handler)

    @property
    def parameters(self):
        return self._pars

    def load(self,
             filepath: str,
             rg_code: Literal["RG_ADD", "RG_REPL", "RG_KEY", "RG_ADDKEY"] = "RG_REPL",
             template: str = None,
             use_template: bool = True):
        """
        Загружает файл данных name в рабочую область в соответствии с шаблоном типа shabl.
        Если поле shabl пусто, то загружается name без шаблона, если пусто поле name, то загружается только шаблон

        :param filepath: Путь до загружаемого файла
        :param rg_code: Режим загрузки при наличии таблицы в рабочей области (RG_ADD, RG_REPL, RG_KEY, RG_ADDKEY, см. комментарий)
        :param template: Путь до шаблона загружаемого файла (Если None и use_template = True, то поиск шаблона выполняется автоматически)
        :param use_template: Использовать шаблон при загрузке файла.
        """

        """
        RG ADD      - Таблица добавляется к имеющейся в рабочей области, совпадение ключевых полей не контролируются 
        (соответствует режиму «Присоединить» в меню)
        RG_REPL     - Таблица в рабочей области замещается (соответствует режиму «Загрузить» в меню)
        RG_KEY      - Данные в таблице, имеющие одинаковые ключевые поля, заменяются. 
        Если ключ не найден, то данные игнорируются (соответствует режиму «Обновить» в меню)
        RG_ADDKEY   - Данные в таблице, имеющие одинаковые ключевые поля, заменяются. 
        Если ключ не найден, то данные вставляются (соответствует режиму «Объединить» в меню)
        """
        match rg_code:
            case "RG_ADD":
                code = 0
            case "RG_REPL":
                code = 1
            case "RG_KEY":
                code = 2
            case "RG_ADDKEY":
                code = 3
            case _:
                raise UnexpectedResult("Некорректный параметр rg_kod!")
        if template is None:
            if use_template:
                self._Astra.Load(code, filepath, self.getTemplate(filepath))
            else:
                self._Astra.Load(code, filepath)
        else:
            self._Astra.Load(code, filepath, template)

    def loadOldFile(self, mode: str, filepath: str, template: str = None) -> None:
        """
        Загрузка файла в рабочую область в старом формате
        :param mode: Режим загрузки файла
        :param filepath: Путь до загружаемого файла
        :param template: Шаблон для загрузки файла
        :return: None
        """
        match mode:
            case "rge":
                mode = 0
            case "cxe":
                mode = 1
            case _:
                raise UnexpectedResult("Некорректный параметр mode!")
        if template is None:
            self._Astra.LoadOldFile(mode, filepath, self.getTemplate(filepath))
        else:
            self._Astra.LoadOldFile(mode, filepath, template)

    def save(self, filepath: str, template: str = None) -> None:
        """
        Сохраняет информацию из рабочей области в файле filepath по шаблону template

        :param filepath: Путь до загружаемого файла
        :param template: Путь до шаблона загружаемого файла
        """
        if template is not None:
            self._Astra.Save(filepath, template)
        else:
            self._Astra.Save(filepath, self.getTemplate(filepath))

    def newFile(self, template: str) -> None:
        """
        Очищается часть рабочей области, соответствующая шаблону.
        Если структура таблицы в рабочей области не существует, она создается (пустая).
        Если таблица уже существует, ее содержимое очищается и приводится в соответствие с шаблоном.
        :param template: Шаблон файла
        """
        self._Astra.NewFile(template)

    create = newFile

    def commit(self) -> None:
        """
        Подтвердить изменения в рабочей области, автоматически выполняется при выполнении Save
        :return: None
        """
        self._Astra.Commit()

    def back(self) -> None:
        """
        Откат, отказ от изменений после последнего Commit
        :return: None
        """
        self._Astra.RollBack()

    def toProtocol(self, text: str) -> None:
        """
        Выводит строку в протокол работы.
        :param text: Текст
        :return:
        """
        self._Astra.Printp(text)

    print = toProtocol

    def jakobi(self, parameters: str = ""):
        return self._Astra.jakobi(parameters)

    def rgm(self, parameters: str = "") -> str:
        """
        Расчет установившегося режима с параметрами params
        :param parameters: Дополнительные параметры расчета. См. ._pars для допустимых значений.
        :return: Результат выполнения расчета (AST_OK - выполнен нормально, AST_NB - не балансируется исходный режим,
        дальнейшее утяжеление невозможно, AST_REPT - утяжеление закончено с заданной точностью или
        достигнуто предельное число итераций)
        """
        """
        Значения params (могут комбинироваться): 
        ""  – c параметрами по умолчанию;
        "p" – расчет с плоского старта;
        "z" – отключение стартового алгоритма;
        "c" – отключение контроля данных;
        "r" – отключение подготовки данных (можно использовать
        при повторном расчете режима с измененными значениями узловых
        мощностей и модулей напряжения). 
        """
        for letter in parameters:
            if letter not in self._pars:
                raise UnexpectedResult(f"Некорректный параметр расчета -> {parameters} -> {letter}")
        result = self._Astra.rgm(parameters)
        return self._calc_codes.get(result)

    ssm = rgm

    def opf(self, parameters: str = "p") -> str:
        """
        Расчет ВРДО с параметрами params
        :param parameters: Дополнительные параметры расчета. См. ._pars для допустимых значений.
        :return: Результат выполнения расчета
        """
        """
        Значения params (могут комбинироваться): 
        ""  – c параметрами по умолчанию;
        "p" – расчет с плоского старта;
        "z" – отключение стартового алгоритма;
        "c" – отключение контроля данных;
        "r" – отключение подготовки данных (можно использовать
        при повторном расчете режима с измененными значениями узловых
        мощностей и модулей напряжения). 
        """
        for letter in parameters:
            if letter not in self._pars:
                raise UnexpectedResult(f"Некорректный параметр расчета -> {parameters} -> {letter}")
        result = self._Astra.opf(parameters)
        return self._calc_codes.get(result)

    def opt(self, parameters: str = "") -> str:
        """
        Оптимизация режима по реактивной мощности
        :param parameters: строка дополнительных параметров
        :return: Результат выполнения расчета
        """
        for letter in parameters:
            if letter not in self._pars:
                raise UnexpectedResult(f"Некорректный параметр расчета -> {parameters} -> {letter}")
        result = self._Astra.opt(parameters)
        return self._calc_codes.get(result)

    def ekv(self, parameters: str = "") -> str:
        """
        Эквивалентирование расчетной модели
        :param parameters: Дополнительные параметры расчета. См. ._pars для допустимых значений.
        :return: Результат выполнения расчета
        """
        """
        Значения params (могут комбинироваться): 
        ""  – c параметрами по умолчанию;
        "p" – расчет с плоского старта;
        "z" – отключение стартового алгоритма;
        "c" – отключение контроля данных;
        "r" – отключение подготовки данных (можно использовать
        при повторном расчете режима с измененными значениями узловых
        мощностей и модулей напряжения). 
        """
        for letter in parameters:
            if letter not in self._pars:
                raise UnexpectedResult(f"Некорректный параметр расчета -> {parameters} -> {letter}")
        result = self._Astra.ekv(parameters)
        return self._calc_codes.get(result)

    def kdd(self, parameters: str = "") -> str:
        """
        Контроль данных для расчета режима
        :param parameters: Дополнительные параметры расчета. См. ._pars для допустимых значений.
        :return: Результат выполнения расчета
        """
        """
        Значения params (могут комбинироваться): 
        ""  – c параметрами по умолчанию;
        "p" – расчет с плоского старта;
        "z" – отключение стартового алгоритма;
        "c" – отключение контроля данных;
        "r" – отключение подготовки данных (можно использовать
        при повторном расчете режима с измененными значениями узловых
        мощностей и модулей напряжения). 
        """
        for letter in parameters:
            if letter not in self._pars:
                raise UnexpectedResult(f"Некорректный параметр расчета -> {parameters} -> {letter}")
        result = self._Astra.kdd(parameters)
        return self._calc_codes.get(result)

    def clearControl(self):
        """
        Инициализировать таблицу значений контролируемых величин
        :return:
        """
        self._Astra.ClearControl()

    def addControl(self, name: str, row_id: int = -1):
        """
        Добавить строку с именем name в таблицу значений контролируемых величин
        :param name: Имя строки
        :param row_id: Индекс строки (по умолчанию -1 - добавит строку в конец таблицы)
        :return:
        """
        self._Astra.AddControl(row_id, name)

    def stepUt(self, parameters: str = "") -> str:
        """
        Выполняет один шаг утяжеления режима по подготовленной траектории.
        Набор параметров совпадает с функцией расчета режима (rgm),
        дополнительно может использоваться параметр "i" – инициализировать значения параметров утяжеления
        (шаг в этом случае не выполняется).
        :param parameters: набор параметров
        :return:
        """
        """
        Возвращает коды:
        AST_OK      – шаг выполнен нормально;
        AST_NB      – не балансируется исходный режим, дальнейшее утяжеление невозможно;
        AST_REPT    – утяжеление закончено с заданной точностью или достигнуто предельное число итераций.
        Понятие "один шаг" зависит от настроек параметров утяжеления:
        если тип утяжеления "Быстрый" – будет проведено полное утяжеление до предела, 
        если "Нормальный" – будет проведено изменение параметров до получения следующего сбалансированного режима
        """
        for letter in parameters:
            if letter not in self._pars:
                raise UnexpectedResult(f"Некорректный параметр расчета -> {parameters} -> {letter}")
        result = self._Astra.step_ut(parameters)
        return self._calc_codes.get(result)

    def ut(self, parameters: str = "") -> str:
        """
        Выполняет автоматическое утяжеления режима по подготовленной траектории.
        Набор параметров совпадает с функцией расчета режима (rgm),
        дополнительно может использоваться параметр "i" – инициализировать значения параметров утяжеления
        (шаг в этом случае не выполняется).
        :param parameters: набор параметров
        :return:
        """
        """
        Возвращает коды:
        AST_OK      – шаг выполнен нормально;
        AST_NB      – не балансируется исходный режим, дальнейшее утяжеление невозможно;
        AST_REPT    – утяжеление закончено с заданной точностью или достигнуто предельное число итераций.
        Понятие "один шаг" зависит от настроек параметров утяжеления:
        если тип утяжеления "Быстрый" – будет проведено полное утяжеление до предела, 
        если "Нормальный" – будет проведено изменение параметров до получения следующего сбалансированного режима
        """
        for letter in parameters:
            if letter not in self._pars:
                raise UnexpectedResult(f"Некорректный параметр расчета -> {parameters} -> {letter}")
        result = self._Astra.ut_utr(parameters)
        return self._calc_codes.get(result)

    weight = ut

    def utFormControl(self):
        """
        Формирует таблицу описаний контролируемых величин, соответствующую заданной траектории утяжеления
        :return:
        """
        self._Astra.ut_FormControl()

    @property
    def lockEvent(self):
        return self._Astra.LockEvent

    @lockEvent.setter
    def lockEvent(self, state: bool):
        """
        Используется для блокирования посылки событий OnChangeData.
        При изменении большого числа данных предварительно устанавливается LockEvent = True (блокировка установлена),
        данные изменяются, затем блокировка снимается lockEvent = False
        и посылается одно сообщение OnChangeData об изменении данных во всей таблице.
        Такой прием позволяет сокращать время операций по изменению данных за счет экономии на обновлении открытых окон.

        :param state: состояние, которое будет установлено для LockEvent
        :return:
        """
        self._Astra.LockEvent = state

    @property
    def renumWP(self):
        return self._Astra.RenumWP

    @renumWP.setter
    def renumWP(self, state: bool):
        """
        Включает (True) или выключает (False) режим изменения ссылок при изменении основного параметра (см. Ссылки)
        :param state: состояние, которое будет установлено для LockEvent
        :return:
        """
        self._Astra.RenumWP = state

    linkUpdate = renumWP

    def utParam(self, parameter: Literal["UT_FORM_P", "UT_ADD_P", "UT_TIP", "UT_STATUS"]):
        """
        Возвращает значение параметров утяжеления (таблица Параметры Утяжеления):

        UT_FORM_P   – формировать описание контролируемых величин (0 – Да, 1 – Нет);
        UT_ADD_P    – добавлять значения в таблицу контролируемых величин;
        UT_TIP      – тип утяжеления (0 – Стандарт, 1 – Быстрый);
        UT_STATUS   – состояние утяжеления (Норма/Предел);

        :param parameter: Параметр утяжеления
        :return: Значение параметра утяжеления
        """
        assert parameter in ["UT_FORM_P", "UT_ADD_P", "UT_TIP", "UT_STATUS"], "Вызван некорректный параметр утяжеления!"

        return self._Astra.ut_Param(parameter)

    @staticmethod
    def getTemplate(file):
        warnings.warn("Метод 'getTemplate' может быть удалена или изменена! "
                      "Используйте другие подходы, описанные в библиотеке pyrastr.", DeprecationWarning, 2)
        extension = os.path.splitext(file)[1]
        templates = os.path.join(os.environ['USERPROFILE'], 'Documents\\RastrWin3\\SHABLON')
        for root, dirs, files in os.walk(templates):
            template = filter(
                lambda x: x if os.path.splitext(x)[1] == extension and 'базовый '
                                                                       'режим мт' not in os.path.splitext(x)[0]
                else None, files)
        return os.path.join(templates, *template)

    def calcIdop(self, temperature: int | float, selection: str = "") -> None:
        """
        Обновление токовых ограничений оборудования на основе ТНВ
        :param temperature: Температура
        :param selection: Выборка строк
        :return:
        """
        self._Astra.CalcIdop(temperature, 0, selection)

    calcIAdditional = calcIdop

    @property
    def isDemo(self):
        return self._Astra.IsDemo

    @property
    def licenseType(self):
        return self._Astra.LicenseType

    @property
    def enableLog(self):
        return self._Astra.LogEnable

    @enableLog.setter
    def enableLog(self, value: bool):
        self._Astra.LogEnable = value


class RastrTables:
    """
    Коллекция таблиц в рабочей области
    """

    def __init__(self, tables):
        """
        Коллекция таблиц в рабочей области
        """
        self._tables = tables

    def __len__(self):
        return self._tables.__len__()

    def __str__(self):
        return self._tables.__str__()

    def addTable(self, name: str):
        """
        Добавляет таблицу name к коллекции, возвращает ее (объект класса RastrTable)
        """
        return RastrTable(self._tables.Add(name))

    add = addTable

    def removeByIndex(self, index: int):
        """
        Удаляет таблицу по указанному индексу
        :param index: Индекс таблицы
        :return:
        """
        self._tables.Remove(index)

    def removeByName(self, name: str):
        """
        Удаляет таблицу по названию таблицы
        :param name: Название таблицы
        :return:
        """
        self._tables.Remove(name)

    def getTableByIndex(self, index: int):
        """
        Возвращает таблицу по указанному индексу
        :param index: Индекс таблицы
        :return: RastrTable
        """
        return RastrTable(self._tables.Item(index))

    def getTableByName(self, name: str):
        """
        Возвращает таблицу по названию
        :param name: Название таблицы
        :return: RastrTable
        """
        return RastrTable(self._tables.Item(name))

    def table(self, item: str | int):
        if isinstance(item, str):
            return self.getTableByName(item)
        elif isinstance(item, int):
            return self.getTableByIndex(item)
        else:
            raise UnexpectedResult("Некорректное наименование или индекс таблицы!")

    t = table
    get = table

    def removeTable(self, item: str | int):
        if isinstance(item, str):
            return self.removeByName(item)
        elif isinstance(item, int):
            return self.removeByIndex(item)
        else:
            raise UnexpectedResult("Некорректное наименование или индекс таблицы!")

    remove = removeTable

    @property
    def count(self) -> int:
        """
        Функция подсчитывает количество таблиц в коллекции в рабочей области
        :return: количество таблиц
        """
        return self._tables.Count

    size = count

    @property
    def list(self) -> list:
        """
        Список таблиц в коллекции в рабочей области
        :return: list
        """
        return list(map(lambda _t: _t.Name, self._tables))


class RowIterator:
    def __init__(self, table, start: int = -1):
        self.idx = start
        self.table = table

    def __iter__(self):
        return self

    def __next__(self):
        self.idx = self.table.FindNextSel(self.idx)
        if self.idx != -1:
            return self.idx
        raise StopIteration



class RastrTable:
    def __init__(self, table):
        self._table = table
        self.columns = RastrColumns(self._table.Cols)
        self._csv_codes = {"CSV_ADD": 0, "CSV_REPL": 1, "CSV_KEY": 2, "CSV_KEYADD": 3, "CSV_REPLNAMES": 5}
        self._cdu_codes = {"CDU_ADD": 0, "CDU_REPL": 1, "CDU_KEY": 2, "CDU_KEYADD": 3}

    def __iter__(self):
        return RowIterator(self._table)

    def get(self,
            row_id: int,
            column: str | int,
            value_type: Literal["scaled", "not_scaled", "scaled_string"] = "not_scaled"):
        c = self.column(column)
        return c.get(row_id, value_type)

    def set(self,
            row_id: int,
            column: str | int,
            value: int | float | str,
            value_type: Literal["scaled", "not_scaled", "scaled_string"] = "not_scaled"):
        c = self.column(column)
        return c.setValue(row_id, value, value_type)

    def addRow(self) -> int:
        """
        Добавляет пустую строку в конец таблицы
        :return: id добавленной строки
        """
        self._table.AddRow()
        return self._table.Size - 1

    add = addRow

    def insertRow(self, row_id: int) -> int:
        """
        Вставляет пустую строку перед строкой с номером row_id
        :param row_id: id строки, перед которой вставляется новая строка
        :return: id вставленной строки
        """
        self._table.InsRow(row_id)
        return row_id - 1

    insert = insertRow

    def duplicateRow(self, row_id: int) -> int:
        """
        Дублирует строку с номером row_id
        :param row_id: id строки, которая будет дублирована
        :return: id новой строки
        """
        self._table.DupRow(row_id)
        return row_id + 1

    duplicate = duplicateRow

    def swapRows(self, row_id_i: int, row_id_j: int):
        """
        Поменять местами две строки в таблице
        :param row_id_i: ROW_ID исходной строки
        :param row_id_j: ROW_ID заменяемой строки
        :return:
        """
        return self._table.swapRow(row_id_i, row_id_j)

    swap = swapRows

    def deleteRow(self, row_id: int):
        """
        Удаляет строку по указанному row_id
        :param row_id: id удаляемой строки
        :return:
        """
        self._table.DelRow(row_id)

    remove = deleteRow
    delete = deleteRow

    def deleteRows(self):
        """
        Удаляет множество строк, определенное текущей выборкой в таблице
        :return:
        """
        self._table.DelRowS()

    removeMany = deleteRows

    def setSelection(self, selection: str) -> int:
        """
        Устанавливает выборку для таблицы
        :param selection: Строка выборки
        :return: Количество строк, попавших в выборку
        """
        self._table.SetSel(selection)
        return self._table.Count

    selection = setSelection

    def clearSelection(self):
        """
        Очищает выборку для таблицы
        :return: Общее количество строк
        """
        return self._table.SetSel("")

    def checkRowSelection(self, row_id: int) -> bool:
        """
        Проверяет входит ли строка с row_id в текущую выборку
        :param row_id: id проверяемой строки
        :return: True - да; False - нет
        """
        return bool(self._table.TestSel(row_id))

    crs = checkRowSelection

    def findNextRowSelection(self, row_id: int = -1) -> int | None:
        """
        Ищет следующую строку от row_id, которая входит в текущую выборку
        :param row_id:
        :return:
        """
        next_row_id = self._table.FindNextSel(row_id)
        if next_row_id == -1:
            return None
        else:
            return next_row_id

    fnrs = findNextRowSelection

    def getRowSelection(self, row_id: int) -> str:
        """
        Формирует выборку, необходимую для точной идентификации строки row_id по ключевым параметрам таблицы
        :param row_id: id передаваемой строки
        :return: строка с выборкой
        """
        return self._table.SelString(row_id)

    grs = getRowSelection

    def writeToCSV(self,
                   filepath: str,
                   parameters: List[str],
                   sep: str = ";",
                   csv_code: Literal["CSV_ADD", "CSV_REPL", "CSV_KEY", "CSV_KEYADD", "CSV_REPLNAMES"] = "CSV_REPL"):
        """
        Записывает часть таблицы в CSV-файл
        :param filepath: Путь к файлу, который будет записан
        :param parameters: Список столбцов, которые будут взяты из таблицы
        :param sep: Разделитель CSV-файла
        :param csv_code: Текстовое значение, определяет режим чтения или записи таблицы.
        Принимает значения: ["CSV_ADD", "CSV_REPL", "CSV_KEY", "CSV_KEYADD", "CSV_REPLNAMES"]
        :return:
        """
        """
        CSV_ADD       - Добавить в конец таблицы (режим «Присоединить»)
        CSV_REPL      - Заменить данные в таблице (режим «Загрузить»)
        CSV_KEY       - Заменить данные по ключевым полям. Если ключ не найден, то данные игнорируются (режим «Обновить»)
        CSV_KEYADD    - Заменить данные по ключевым полям. Если ключ не найден, то данные добавляются в конец таблицы (режим «Объединить»)
        CSV_REPLNAMES - В отличие от CSV_REPL игнорируется первая строка в файле, полагая, что в ней содержатся названия полей
        """
        try:
            self._table.WriteCSV(self._csv_codes.get(csv_code), filepath, ",".join(parameters), sep)
        except Exception as err:
            raise UnexpectedResult(err)

    writeCSV = writeToCSV

    def readCSV(self, filepath: str,
                parameters: List[str],
                sep: str = ";",
                csv_code: Literal["CSV_ADD", "CSV_REPL", "CSV_KEY", "CSV_KEYADD", "CSV_REPLNAMES"] = "CSV_REPL",
                default_parameters: str = ""):
        """
        Читает CSV-файл
        :param filepath: Путь к файлу, который будет прочитан
        :param parameters: Список столбцов, которые будут взяты из файла
        :param sep: Разделитель CSV-файла
        :param csv_code: Числовое значение, определяет режим чтения или записи таблицы.
        Принимает значения: ["CSV_ADD", "CSV_REPL", "CSV_KEY", "CSV_KEYADD", "CSV_REPLNAMES"]
        :param default_parameters: Значения по умолчанию в виде «имя=значение»
        :return:
        """
        """
        CSV_ADD       - Добавить в конец таблицы (режим «Присоединить»)
        CSV_REPL      - Заменить данные в таблице (режим «Загрузить»)
        CSV_KEY       - Заменить данные по ключевым полям. Если ключ не найден, то данные игнорируются (режим «Обновить»)
        CSV_KEYADD    - Заменить данные по ключевым полям. Если ключ не найден, то данные добавляются в конец таблицы (режим «Объединить»)
        CSV_REPLNAMES - В отличие от CSV_REPL игнорируется первая строка в файле, полагая, что в ней содержатся названия полей
        """
        try:
            self._table.ReadCSV(self._csv_codes.get(csv_code), filepath, ",".join(parameters), sep, default_parameters)
        except Exception as err:
            raise UnexpectedResult(err)

    def writeToCDU(self,
                   filepath: str,
                   parameters: List[str],
                   cdu_code: Literal["CDU_ADD", "CDU_REPL", "CDU_KEY", "CDU_KEYADD"] = "CDU_ADD"):
        """
        Записывает информацию из таблицы в файл, по структуре соответствующий формату ЦДУ.
        В начало каждой строки записывается четырехзначный код.
        Затем следует одно 4-позиционное поле и произвольное число 8-позиционных. Для пропуска поля ставится знак «$»
        :param filepath: Путь к файлу, который будет записан
        :param parameters: Список столбцов, которые будут взяты из таблицы
        :param cdu_code: Числовое значение, определяет режим чтения или записи таблицы.
        Принимает значения: ["CDU_ADD", "CDU_REPL" "CDU_KEY", "CDU_KEYADD"]
        :return:
        """
        """
        CDU_ADD       - Добавить в конец таблицы (режим «Присоединить»)
        CDU_REPL      - Заменить данные в таблице (режим «Загрузить»)
        CDU_KEY       - Заменить данные по ключевым полям. Если ключ не найден, то данные игнорируются (режим «Обновить»)
        CDU_KEYADD    - Заменить данные по ключевым полям. Если ключ не найден, то данные добавляются в конец таблицы (режим «Объединить»)
        """
        try:
            self._table.WriteCDU(self._cdu_codes.get(cdu_code), filepath, ",".join(parameters))
        except Exception as err:
            raise UnexpectedResult(err)

    writeCDU = writeToCDU

    def readCDU(self,
                filepath: str, parameters: List[str],
                cdu_code: Literal["CDU_ADD", "CDU_REPL", "CDU_KEY", "CDU_KEYADD"] = "CDU_REPL",
                default_parameters: str = ""):
        """
        Читает файл в формате ЦДУ. При чтении файла в макете ЦДУ происходит выборка из файла строк, соответствующих
        заданному коду. Дальнейшая обработка аналогична CSV-файлу
        :param filepath: Путь к файлу, который будет записан
        :param parameters: Список столбцов, которые будут взяты из таблицы
        :param cdu_code: Числовое значение, определяет режим чтения или записи таблицы.
        Принимает значения: ["CDU_ADD", "CDU_REPL" "CDU_KEY", "CDU_KEYADD"]
        :param default_parameters: Значения по умолчанию в виде «имя=значение»
        :return:
        """
        """
        CDU_ADD       - Добавить в конец таблицы (режим «Присоединить»)
        CDU_REPL      - Заменить данные в таблице (режим «Загрузить»)
        CDU_KEY       - Заменить данные по ключевым полям. Если ключ не найден, то данные игнорируются (режим «Обновить»)
        CDU_KEYADD    - Заменить данные по ключевым полям. Если ключ не найден, то данные добавляются в конец таблицы (режим «Объединить»)
        """
        try:
            self._table.ReadCDU(self._cdu_codes.get(cdu_code), filepath, ",".join(parameters), default_parameters)
        except Exception as err:
            raise UnexpectedResult(err)

    def column(self, item: str | int):
        if isinstance(item, str):
            return self.columns.getByName(item)
        elif isinstance(item, int):
            return self.columns.getByIndex(item)
        else:
            raise UnexpectedResult("Некорректное наименование или индекс столбца")

    c = column

    def iterRows(self):
        """
        Формирует итератор, который будет возвращать row_id следующей строки в текущей выборке
        :return:
        """
        row_id = self.findNextRowSelection()
        while row_id is not None:
            yield row_id
            row_id = self.findNextRowSelection(row_id)

    iterr = iterRows

    def iterColumns(self):
        for column_id in range(0, self.columns.count):
            column = self.column(column_id)
            yield column
            column_id += 1

    iterc = iterColumns

    @property
    def count(self) -> int:
        """
        Подсчитывает число строк в текущей выборке
        :return: число строк в текущей выборке
        """
        return self._table.Count

    size = count

    @property
    def name(self) -> str:
        """
        Возвращает имя таблицы
        :return: имя таблицы
        """
        return self._table.Name

    @name.setter
    def name(self, name_: str):
        """
        Устанавливает имя таблицы
        """
        self._table.Name = name_

    @property
    def description(self) -> str:
        """
        Возвращает описание таблицы
        :return: описание таблицы
        """
        return self._table.Description

    @description.setter
    def description(self, description_: str):
        """
        Устанавливает описание таблицы
        """
        self._table.Description = description_

    @property
    def keys(self):
        return self._table.Key

    @keys.setter
    def keys(self, keys_: List[str]):
        self._table.Key = ",".join(keys_)

    @property
    def template(self) -> str:
        """
        Возвращает имя шаблона (типа файла), в который помещается данная таблица
        :return: имя шаблона
        """
        return self._table.TemplateName

    @template.setter
    def template(self, template_name: str):
        """
        Устанавливает имя шаблона (типа файла), в который помещается данная таблица
        """
        self._table.TemplateName = template_name

    @property
    def rowsCount(self) -> int:
        """
        Подсчитывает количество строк в таблице
        :return: число строк в таблице
        """
        return self._table.Size

    fullSize = rowsCount


class RastrColumns:
    """
    Коллекция столбцов в RastrTable
    """

    def __init__(self, cols):
        self._cols = cols
        self._column_types = ["PR_INT", "PR_REAL", "PR_STRING", "PR_BOOL", "PR_ENUM", "PR_ENPIC", "PR_COLOR"]

    def getByIndex(self, index: int):
        """
        Возвращает столбец таблицы RastrTable по индексу
        :param index: индекс столбца
        :return: RastrColumn
        """
        return RastrColumn(self._cols.Item(index))

    def getByName(self, name: str):
        """
        Возвращает столбец таблицы RastrTable по наименованию
        :param name: наименование столбца
        :return: RastrColumn
        """
        return RastrColumn(self._cols.Item(name))

    def removeByIndex(self, index: int):
        """
        Удаляет столбец из таблицы по установленному индексу
        :param index: индекс столбца
        :return:
        """
        self._cols.Remove(index)

    def removeByName(self, name: str):
        """
        Удаляет столбец из таблицы по наименованию
        :param name: наименование столбца
        :return:
        """
        self._cols.Remove(name)

    def add(self, name: str,
            column_type: Literal["PR_INT", "PR_REAL", "PR_STRING", "PR_BOOL", "PR_ENUM", "PR_ENPIC", "PR_COLOR"]):
        """
        Добавляет столбец в таблице с установленным наименованием и типом данных
        :param name: наименование столбца
        :param column_type: тип данных,
        ["PR_INT", "PR_REAL", "PR_STRING", "PR_BOOL", "PR_ENUM", "PR_ENPIC", "PR_COLOR"]. См. инструкцию
        :return:
        """
        assert column_type in self._column_types, "Некорректный тип значения столбца!"

        return RastrColumn(self._cols.Add(name, self._column_types.index(column_type)))

    def find(self, name: str):
        return self._cols.Find(name)

    @property
    def count(self):
        """
        Подсчитывает число столбцов в коллекции
        :return:
        """
        return self._cols.Count


class RastrColumn:
    """
    Столбец таблицы из коллекции RastrColumns
    """

    def __init__(self, column):
        """
        Столбец из коллекции RastrColumns
        """
        self._column = column
        self._property_types = ["FL_NAME", "FL_TIP", "FL_WIDTH", "FL_PREC",
                                "FL_ZAG", "FL_FORMULA", "FL_AFOR", "FL_XRM",
                                "FL_NAMEREF", "FL_DESC", "FL_MIN", "FL_MAX", "FL_MASH"]
        self._value_types = ["scaled", "not_scaled", "scaled_string"]

    def calc(self, formula: str):
        """
        Вычисляет значения элементов столбца в соответствии с заданной формулой и выборкой (групповая коррекция).
        :param formula:
        :return:
        """
        self._column.Calc(formula)

    def getProperty(self,
                    property_type: Literal["FL_NAME", "FL_TIP", "FL_WIDTH", "FL_PREC",
                    "FL_ZAG", "FL_FORMULA", "FL_AFOR", "FL_XRM",
                    "FL_NAMEREF", "FL_DESC", "FL_MIN", "FL_MAX", "FL_MASH"]):
        """
        Получает значение свойства столбца
        :param property_type: свойство столбца. Может быть ["FL_NAME", "FL_TIP", "FL_WIDTH", "FL_PREC",
                                "FL_ZAG", "FL_FORMULA", "FL_AFOR", "FL_XRM",
                                "FL_NAMEREF", "FL_DESC", "FL_MIN", "FL_MAX", "FL_MASH"]. См. Инструкцию
        :return: значение свойства
        """
        assert property_type in self._property_types, "Некорректный тип свойства столбца!"

        return self._column.Prop(self._property_types.index(property_type))

    def setProperty(self,
                    property_type: Literal["FL_NAME", "FL_TIP", "FL_WIDTH", "FL_PREC",
                    "FL_ZAG", "FL_FORMULA", "FL_AFOR", "FL_XRM",
                    "FL_NAMEREF", "FL_DESC", "FL_MIN", "FL_MAX", "FL_MASH"],
                    value):
        """
        Устанавливает значение свойства столбца
        :param property_type: свойство столбца. Может быть ["FL_NAME", "FL_TIP", "FL_WIDTH", "FL_PREC",
                                "FL_ZAG", "FL_FORMULA", "FL_AFOR", "FL_XRM",
                                "FL_NAMEREF", "FL_DESC", "FL_MIN", "FL_MAX", "FL_MASH"]. См. Инструкцию
        :param value: значение для свойства
        """
        assert property_type in self._property_types, "Некорректный тип свойства столбца!"

        self._column.SetProp(self._property_types.index(property_type), value)

    @property
    def name(self) -> str:
        """
        Возвращает имя столбца
        :return:
        """
        return self._column.Name

    def setValue(self,
                 row_id: int,
                 value: int | float | str,
                 value_type: Literal["scaled", "not_scaled", "scaled_string"] = "not_scaled"):
        """
        Устанавливает значение элемента в строке row_id
        :param row_id: id строки, в которой будет установлено значение
        :param value: значение, которое будет установлено
        :param value_type: Тип устанавливаемого значения.
        Может принимать значения: ["scaled", "not_scaled", "scaled_string"]. См. Инструкцию
        :return:
        """
        assert value_type in self._value_types, f"Некорректный тип величины! " \
                                                f"Тип может принимать значения: [{self._value_types}]."
        match value_type:
            case "scaled":
                self._column.SetZN(row_id, value)
            case "not_scaled":
                self._column.SetZ(row_id, value)
            case "scaled_string":
                self._column.SetZS(row_id, value)

    def getValue(self,
                 row_id: int,
                 value_type: Literal["scaled", "not_scaled", "scaled_string"] = "not_scaled"):
        """
        Возвращает значение элемента в строке row_id
        :param row_id: id строки, в которой будет прочитано значение
        :param value_type: Тип возвращаемого значения.
        Может принимать значения: ["scaled", "not_scaled", "scaled_string"]. См. Инструкцию
        :return:
        """
        if value_type not in self._value_types:
            raise Exception(f"Некорректный тип величины! Тип может принимать значения: [{self._value_types}].")
        match value_type:
            case "scaled":
                return self._column.ZN(row_id)
            case "not_scaled":
                return self._column.Z(row_id)
            case "scaled_string":
                return self._column.ZS(row_id)

    get = getValue

    set = setValue


class UnexpectedResult(Exception):
    def __init__(self, *args):
        if args:
            self.message = args[0]
        else:
            self.message = None

    def __str__(self):
        if self.message:
            return self.message
        else:
            return 'Неожиданный результат выполнения!'

