/*******************************************************************
 *
 *               Макрос Построитель отчетов работы интернет-магазина
 *                                 (точка входа)
 *
 ********************************************************************/

/**
 * Основная точка входа макроса для построения отчетов по работе интернет-магазина.
 * Выполняется сразу при запуске и реализует следующую последовательность действий:
 *
 * 1. Загружает пользовательские настройки (путь к CSV-файлу и период анализа)
 * 2. Преобразует путь к файлу в корректный URL формат
 * 3. Загружает данные из CSV-файла
 * 4. Разбирает CSV в структуру данных для дальнейшего анализа
 * 5. Создает фильтр по периоду на основе пользовательских настроек
 * 6. Анализирует данные для получения статистики по категориям и брендам
 * 7. Создает два отчета с графиками, используя универсальный инструмент:
 *    - Отчет по категориям товаров
 *    - Отчет по брендам
 * 8. Уведомляет пользователя о завершении работы
 *
 * В случае возникновения ошибок на любом этапе выводит сообщение пользователю.
 *
 * @async
 * @function
 * @returns {void}
 */
(async function () {
  try {
    // Этап 1: Получение пользовательских настроек из листа "Анализ данных"
    /** @type {UserSettings} */
    const settings = readUserSettings();

    // Конвертируем путь к файлу в формат URL для корректной загрузки
    settings.url = convertPath(settings.url);
    console.info('settings', settings);

    // Этап 2: Загрузка данных из CSV по указанному пути
    console.info('Получение данных');
    const xhr = await loadData(settings.url);
    const csvString = xhr.responseText;

    // Этап 3: Парсинг CSV-данных в структурированный формат
    /** @type {ShopEventEntry[]} */
    const data = parseCsv(csvString, ';');

    console.info('Полученные данные', data);

    // Этап 4: Создание фильтра для анализа по указанному периоду
    /**@type {AnalitycsFilter} */
    const filter = {
      period: {
        start: settings.startDate,
        end: settings.endDate,
      },
    };

    // Этап 5: Анализ данных и формирование статистики по категориям и брендам
    const analitics = checkAndAnalyzeData(data, filter);

    // Этап 6: Создание отчетов с графиками по категориям и брендам
    // Параметры 3, 1 задают отступ от левого верхнего угла листа
    createPurchaseReport(analitics.purchasePerCategories, 'Покупки по категориям', 3, 1);
    createPurchaseReport(analitics.purchasePerBrands, 'Покупки по брендам', 3, 1);

    // Этап 7: Уведомление пользователя о завершении работы
    uiAlert('Сообщение', 'Отчет готов!');

    // Этап 8: Обновляем представление для корректного отображения созданных графиков и данных
    /**
     * Обновляет визуальное представление листа и всех графиков
     *
     * Вызов API метода Asc.editor.controller.view.resize() запускает принудительную
     * перерисовку всего документа, что необходимо для корректного отображения созданных
     * программно графиков и обновленных данных. Без этого вызова графики могут быть
     * отображены некорректно или вообще не отображены до следующего действия пользователя.
     *
     * @private
     */
    Asc.editor.controller.view.resize();
  } catch (error) {
    // Логируем ошибку в консоль для отладки
    console.error('Ошибка макроса:', error);

    // Выводим пользователю понятное сообщение об ошибке
    uiAlert('Ошибка', error.message);
  }
})();

/*******************************************************************
 *
 *                            Основной модуль
 *
 ********************************************************************/

/**
 * Пользовательские настройки
 *
 * @typedef {{
 *   url: string,
 *   startDate: Date,
 *   endDate: Date
 * }} UserSettings
 */

/**
 * Читает пользовательские настройки из листа "Анализ данных"
 *
 * @description
 * Функция получает настройки из листа "Анализ данных":
 * - URL путь к CSV-файлу из ячейки C4
 * - Дату начала периода анализа из ячейки C6
 * - Дату окончания периода анализа из ячейки C7
 *
 * Даты конвертируются из формата R7 Office в объекты JavaScript Date.
 * При отсутствии дат используются значения по умолчанию.
 *
 * @returns {UserSettings} Объект с настройками пользователя
 * @throws {Error} Если лист "Анализ данных" не найден
 *
 * @example
 * const settings = readUserSettings();
 * console.log(settings.url); // URL для загрузки данных
 * console.log(settings.startDate); // Дата начала периода анализа
 */
function readUserSettings() {
  // Имя листа, из которого будем получать настройки
  const sheetName = 'Анализ данных';

  // Получаем лист по имени
  const sheet = Api.GetSheet(sheetName);
  if (sheet === null) {
    // Если лист не найден, выбрасываем исключение
    throw new Error(`Не найден лист "${sheetName}"!`);
  }

  // Получаем путь к CSV-файлу из ячейки C4
  const urlRange = sheet.GetRange('C4');
  const url = urlRange.GetValue();

  // Получаем дату начала периода из ячейки C6
  const startDateRange = sheet.GetRange('C6');
  let startDate = startDateRange.GetValue();
  // Если дата не указана или некорректна, используем значение по умолчанию (1 января 2000 года)
  startDate =
    startDate === null || startDate === '' || !Number(startDate) ? new Date(2000, 0, 1) : r7SerialToJsDate(startDate);

  // Получаем дату окончания периода из ячейки C7
  const endDateRange = sheet.GetRange('C7');
  let endDate = endDateRange.GetValue();
  // Если дата не указана или некорректна, используем значение по умолчанию (31 декабря 2100 года)
  endDate =
    endDate === null || endDate === '' || !Number(endDate)
      ? new Date(2100, 0, 0, 23, 59, 59, 99)
      : r7SerialToJsDate(endDate);

  // Возвращаем объект с настройками
  return { url, startDate, endDate };
}

/**
 * @typedef {{
 *  brand: string,
 *  category_code: string,
 *  event_time: string,
 *  event_type: string,
 *  price: string,
 *  product_id: string,
 *  user_id: string,
 *  user_session: string
 * }} ShopEventEntry
 */

/**
 * @typedef {{
 *  brand: string,
 *  category_code: string | null,
 *  category_code_lv0: string,
 *  category_code_lv1: string,
 *  category_code_lv2: string,
 *  event_time: Date,
 *  event_type: 'view' | 'cart' | 'purchase',
 *  price: number,
 *  product_id: string,
 *  user_id: string,
 *  user_session: string
 * }} ShopEvent
 */

/**
 * @typedef {{
 *   start?: Date,
 *   end?: Date
 * }} AnalitycsFilterPeriod
 */

/**
 * @typedef {{
 *   period?: AnalitycsFilterPeriod
 * }} AnalitycsFilter
 */

/**
 * @typedef {{
 *   count: number, // Число покупок
 *   price: number, // Сумма покупок
 *   avgPrice: number // Средняя цена одной покупки
 * }} PurchasePerCategory
 */

/**
 * @typedef {{
 *   count: number, // Число покупок
 *   price: number, // Сумма покупок
 *   avgPrice: number // Средняя цена одной покупки
 * }} PurchasePerBrand
 */

/**
 * @typedef {Object.<string, PurchasePerCategory>} PurchasePerCategories
 */

/**
 * @typedef {Object.<string, PurchasePerBrand>} PurchasePerBrands
 */

/**
 * @typedef {{
 *  purchasePerCategories: PurchasePerCategories,
 *  purchasePerBrands: PurchasePerBrands
 * }} Analitycs
 */

/**
 * Проверяет и анализирует данные магазина для формирования статистики покупок
 *
 * @description
 * Функция принимает данные о событиях в магазине (просмотры, добавления в корзину, покупки)
 * и анализирует их для получения статистики по категориям товаров и брендам.
 *
 * Вычисляются следующие метрики:
 * - Количество покупок по категориям и брендам
 * - Сумма покупок по категориям и брендам
 * - Средний чек покупок по категориям и брендам
 *
 * Функция также поддерживает фильтрацию данных по периоду времени.
 *
 * @param {ShopEventEntry[]} data - Массив объектов с данными о событиях магазина
 * @param {AnalitycsFilter} [filter] - Фильтр для анализа данных (например, по периоду времени)
 * @returns {Analitycs} - Результат анализа данных с метриками по категориям и брендам
 * @throws {Error} Если переданный аргумент data не является массивом
 * @throws {Error} Если данные не соответствуют ожидаемому формату ShopEventEntry
 *
 * @example
 * // Базовый пример использования
 * const result = checkAndAnalyzeData(shopData);
 * console.log(result.purchasePerCategories); // Статистика по категориям
 *
 * @example
 * // С фильтрацией по периоду
 * const filter = {
 *   period: {
 *     start: new Date('2023-01-01'),
 *     end: new Date('2023-12-31')
 *   }
 * };
 * const result = checkAndAnalyzeData(shopData, filter);
 */
function checkAndAnalyzeData(data, filter) {
  // Проверка, что data является массивом
  if (!Array.isArray(data)) {
    throw new Error('Аргумент data должен быть массивом объектов типа ShopEventEntry');
  }

  // Проверка элементов массива на соответствие типу ShopEventEntry
  // Проверяем только наличие обязательных полей, дополнительные поля игнорируются
  if (data.length > 0) {
    // Список обязательных свойств для типа ShopEventEntry
    const requiredProps = [
      'brand',
      'category_code',
      'event_time',
      'event_type',
      'price',
      'product_id',
      'user_id',
      'user_session',
    ];
    const missingProps = [];

    // Проверяем до 3-х первых элементов для надежности валидации
    for (const item of data.slice(0, 3)) {
      for (const prop of requiredProps) {
        if (!(prop in item) && !missingProps.includes(prop)) {
          missingProps.push(prop);
        }
      }

      // Если нашли все свойства, выходим из цикла
      if (missingProps.length === 0) {
        break;
      }
    }

    // Выдаем ошибку только если есть отсутствующие обязательные свойства
    if (missingProps.length > 0) {
      throw new Error(`Данные не соответствуют типу ShopEventEntry: отсутствуют свойства ${missingProps.join(', ')}`);
    }
  }

  // Инициализация объектов для сбора статистики по категориям и брендам
  const purchasePerCategories = {};
  const purchasePerBrands = {};

  // Обрабатываем каждый элемент из входных данных
  for (const item of data) {
    // Преобразуем строковое представление времени в объект Date
    const eventTime = new Date(item.event_time);

    // Фильтрация по начальной дате, если указана
    if (filter?.period?.start && filter.period.start.getTime) {
      if (filter.period.start.getTime() > eventTime.getTime()) {
        continue; // Пропускаем события до начальной даты
      }
    }

    // Фильтрация по конечной дате, если указана
    if (filter?.period?.end && filter.period.end.getTime) {
      if (filter.period.end.getTime() < eventTime.getTime()) {
        continue; // Пропускаем события после конечной даты
      }
    }

    // Разбиваем иерархическую категорию на уровни, используя значения по умолчанию '_none'
    const [category_code_lv0 = '_none', category_code_lv1 = '_none', category_code_lv2 = '_none'] = item.category_code
      ? item.category_code.split('.')
      : [];

    // Преобразуем строковое представление цены в число, удаляя пробелы и заменяя запятую на точку
    const price = Number(String(item.price).replace(/\s/g, '').replace(',', '.'));

    // Создаем объект события с нормализованными данными
    const event = {
      brand: item.brand || '_none', // Если бренд не указан, используем '_none'
      category_code: item.category_code,
      category_code_lv0,
      category_code_lv1,
      category_code_lv2,
      event_time: eventTime,
      event_type: item.event_type,
      price,
      product_id: item.product_id,
      user_id: item.user_id,
      user_session: item.user_session,
    };

    // Инициализация статистики для категории, если она встречается впервые
    if (!purchasePerCategories[category_code_lv0]) {
      purchasePerCategories[category_code_lv0] = {
        count: 0, // Количество покупок
        price: 0, // Общая сумма покупок
        avgPrice: 0, // Средняя стоимость покупки
      };
    }

    // Инициализация статистики для бренда, если он встречается впервые
    if (!purchasePerBrands[event.brand]) {
      purchasePerBrands[event.brand] = {
        count: 0,
        price: 0,
        avgPrice: 0,
      };
    }

    // Учитываем только события с типом 'purchase'
    if (event.event_type === 'purchase') {
      // Обновляем статистику по категории
      purchasePerCategories[category_code_lv0].count += 1;
      purchasePerCategories[category_code_lv0].price += event.price;
      // Пересчитываем среднюю цену после каждой новой покупки
      purchasePerCategories[category_code_lv0].avgPrice =
        purchasePerCategories[category_code_lv0].price / purchasePerCategories[category_code_lv0].count;

      // Обновляем статистику по бренду
      purchasePerBrands[event.brand].count += 1;
      purchasePerBrands[event.brand].price += event.price;
      purchasePerBrands[event.brand].avgPrice =
        purchasePerBrands[event.brand].price / purchasePerBrands[event.brand].count;
    }
  }

  // Возвращаем объект с результатами анализа
  return { purchasePerCategories, purchasePerBrands };
}

/**
 * Создает отчет о покупках по категориям или брендам с визуализацией в виде графиков и таблиц
 *
 * @param {PurchasePerCategories|PurchasePerBrands} data - Объект со статистикой покупок
 *                                                        (по категориям или по брендам)
 * @param {string} title - Название листа для отчета
 * @param {number} [firstRowIndex=0] - Индекс первой строки для размещения отчета (нумерация с 0)
 * @param {number} [firstColumnIndex=0] - Индекс первого столбца для размещения отчета (нумерация с 0)
 *
 * @description
 * Функция выполняет следующие действия:
 * 1. Создает новый лист с указанным названием (удаляя старый, если он существует)
 * 2. Формирует заголовок отчета
 * 3. Создает секцию "Самые популярные" (топ-3 по количеству покупок)
 * 4. Создает секцию "Самые дорогие чеки" (топ-3 с наибольшей средней ценой покупки)
 * 5. Формирует таблицу со статистикой:
 *    - Количество покупок (до 15 записей, отсортированных по алфавиту)
 *    - Средняя цена покупки (до 15 записей, отсортированных по алфавиту)
 * 6. Создает два столбчатых графика на основе таблиц статистики
 *
 * @throws {Error} Если листа с указанным названием нет в документе и его не удается создать
 * @returns {void}
 *
 * @example
 * // Создать отчет по категориям на новом листе
 * createPurchaseReport(analitics.purchasePerCategories, "Покупки по категориям");
 *
 * // Создать отчет по брендам со смещением от верхнего левого угла
 * createPurchaseReport(analitics.purchasePerBrands, "Покупки по брендам", 3, 1);
 */
function createPurchaseReport(data, title, firstRowIndex = 0, firstColumnIndex = 0) {
  // Удаляем существующий лист с таким названием, если он есть
  let sheet = Api.GetSheet(title);
  if (sheet) {
    sheet.Delete();
  }

  // Создаем новый лист и делаем его активным
  Api.AddSheet(title);
  sheet = Api.GetActiveSheet();

  // Используем firstRow для совместимости с 1-based индексацией ячеек в R7 Office
  const firstRow = firstRowIndex + 1;

  // Создаем и форматируем заголовок отчета
  const titleRange = sheet.GetRangeByNumber(firstRowIndex, firstColumnIndex + 2);
  titleRange.SetFontSize(16);
  titleRange.SetBold(true);
  titleRange.SetValue(title);

  // Устанавливаем оптимальную ширину основных колонок для улучшения читаемости
  sheet.SetColumnWidth(firstColumnIndex, 20); // Колонка с заголовками секций
  sheet.SetColumnWidth(firstColumnIndex + 1, 20); // Колонка с названиями категорий (левая таблица)
  sheet.SetColumnWidth(firstColumnIndex + 9, 20); // Колонка с названиями категорий (правая таблица)
  sheet.SetColumnWidth(firstColumnIndex + 10, 20); // Колонка с числовыми значениями (правая таблица)

  // Получаем список всех категорий из объекта данных
  const keys = Object.keys(data);

  // Создаем заголовок первой секции "Самые популярные группы" (категории или бренды)
  const bestSellsTitleRange = sheet.GetRangeByNumber(firstRowIndex + 2, firstColumnIndex);
  bestSellsTitleRange.SetFontSize(12);
  bestSellsTitleRange.SetBold(true);
  bestSellsTitleRange.SetValue('Самые популярные');

  // Выводим топ-3 группы по количеству покупок (сортировка по убыванию count)
  keys
    .sort((a, b) => data[b].count - data[a].count) // Сортировка по убыванию количества покупок
    .slice(0, 3) // Берем только первые 3 записи
    .forEach((key, i) => {
      // Добавляем название категории и количество покупок в соответствующие ячейки
      sheet.GetCells(i + firstRow + 3, firstColumnIndex + 1).SetValue(key);
      sheet.GetCells(i + firstRow + 3, firstColumnIndex + 2).SetValue(data[key].count);
    });

  // Создаем заголовок для раздела "Самые дорогие чеки"
  const bestCheckTitleRange = sheet.GetRangeByNumber(firstRowIndex + 7, firstColumnIndex);
  bestCheckTitleRange.SetFontSize(12);
  bestCheckTitleRange.SetBold(true);
  bestCheckTitleRange.SetValue('Самые дорогие чеки');

  // Выводим топ-3 группы по средней стоимости покупки (сортировка по убыванию avgPrice)
  keys
    .sort((a, b) => data[b].avgPrice - data[a].avgPrice) // Сортировка по убыванию средней цены
    .slice(0, 3) // Берем только первые 3 записи
    .forEach((key, i) => {
      // Добавляем название категории
      sheet.GetCells(i + firstRow + 8, firstColumnIndex + 1).SetValue(key);

      // Форматируем значение средней цены (замена точки на запятую для отображения в российском формате)
      const avgPriceStr = String(data[key].avgPrice).replace('.', ',');
      sheet.GetCells(i + firstRow + 8, firstColumnIndex + 2).SetValue(avgPriceStr);

      // Устанавливаем денежный формат с двумя знаками после запятой
      sheet.GetCells(i + firstRow + 8, firstColumnIndex + 2).SetNumberFormat('#,##0.00');
    });

  // Определяем начальную строку для размещения основных таблиц
  // Оставляем 12 строк от начала для предыдущих разделов отчета
  const firstRowValues = firstRow + 12;

  // Создаем и оформляем заголовок таблицы с количеством покупок (левая таблица)
  const countPurchaseRange = getRangeBySize(sheet, firstRowValues, firstColumnIndex, 1, 2);
  countPurchaseRange.SetAlignHorizontal('center');
  countPurchaseRange.SetFillColor(Api.CreateColorFromRGB(91, 149, 249)); // Синий цвет для заголовка
  countPurchaseRange.SetValue(['Группа', 'Количество покупок']);

  // Создаем и оформляем заголовок таблицы со средней ценой покупки (правая таблица)
  const avgPurchaseRange = getRangeBySize(sheet, firstRowValues, firstColumnIndex + 9, 1, 2);
  avgPurchaseRange.SetAlignHorizontal('center');
  avgPurchaseRange.SetFillColor(Api.CreateColorFromRGB(244, 101, 36)); // Оранжевый цвет для заголовка
  avgPurchaseRange.SetValue(['Группа', 'Средняя цена покупки']);
  // Заполняем таблицу с количеством покупок (левая таблица)
  keys
    .sort((a, b) => data[b].count - data[a].count) // Сначала сортируем по убыванию количества
    .slice(0, 15) // Ограничиваем список 15-ю записями
    .sort((a, b) => String(a).localeCompare(String(b))) // Сортируем по алфавиту для лучшей навигации
    .forEach((key, i) => {
      const { count } = data[key];

      // Добавляем название категории и количество покупок
      sheet.GetCells(i + firstRowValues + 2, firstColumnIndex + 1).SetValue(key);
      sheet.GetCells(i + firstRowValues + 2, firstColumnIndex + 2).SetValue(count);
    });

  // Заполняем таблицу со средней ценой покупки (правая таблица)
  keys
    .sort((a, b) => data[b].avgPrice - data[a].avgPrice) // Сначала сортируем по убыванию средней цены
    .slice(0, 15) // Ограничиваем список 15-ю записями
    .sort((a, b) => String(a).localeCompare(String(b))) // Сортируем по алфавиту для лучшей навигации
    .forEach((key, i) => {
      const { avgPrice } = data[key];
      // Форматируем значение средней цены (замена точки на запятую для отображения в российском формате)
      const avgPriceStr = String(avgPrice).replace('.', ',');

      // Добавляем название категории и среднюю цену покупки
      sheet.GetCells(i + firstRowValues + 2, firstColumnIndex + 10).SetValue(key);
      sheet.GetCells(i + firstRowValues + 2, firstColumnIndex + 11).SetValue(avgPriceStr);
      // Устанавливаем денежный формат с двумя знаками после запятой
      sheet.GetCells(i + firstRowValues + 2, firstColumnIndex + 11).SetNumberFormat('#,##0.00');
    });

  // Получаем адрес диапазона данных для первого графика (включает заголовок + до 15 строк данных)
  const countPurchaseValuesAddress = getRangeBySize(sheet, firstRowValues, firstColumnIndex, 15 + 1, 2).Address;

  // Создаем столбчатый график для визуализации количества покупок
  Api.GetActiveSheet().AddChart(
    `'${title}'!${countPurchaseValuesAddress}`, // Ссылка на диапазон данных
    true, // С заголовками
    'bar', // Тип графика - столбчатый
    2, // Стиль графика
    105 * 36000, // Ширина в тваипах (1 тваип = 1/36000 дюйма)
    105 * 36000, // Высота в тваипах
    firstColumnIndex + 2, // Отступ от левого края, в тваипах
    2 * 36000, // Отступ от верхнего края, в тваипах
    firstRowValues, // Позиция по вертикали в тваипах
    3 * 36000 // Глубина (для 3D графиков)
  );

  // Получаем адрес диапазона данных для второго графика (включает заголовок + до 15 строк данных)
  const avgPurchaseValuesAddress = getRangeBySize(sheet, firstRowValues, firstColumnIndex + 9, 15 + 1, 2).Address;

  // Создаем столбчатый график для визуализации средней цены покупки
  Api.GetActiveSheet().AddChart(
    `'${title}'!${avgPurchaseValuesAddress}`, // Ссылка на диапазон данных
    true, // С заголовками
    'bar', // Тип графика - столбчатый
    2, // Стиль графика
    105 * 36000, // Ширина в тваипах
    105 * 36000, // Высота в тваипах
    firstColumnIndex + 11, // Отступ от левого края
    2 * 36000, // Отступ от верхнего края
    firstRowValues, // Позиция по вертикали
    3 * 36000 // Глубина (для 3D графиков)
  );
}

/*******************************************************************
 *
 *                            Инструменты
 *
 ********************************************************************/

/**
 * Загружает данные по указанному URL и возвращает Promise с результатом
 *
 * @description
 * Функция выполняет асинхронный HTTP-запрос для загрузки данных.
 * При успешной загрузке Promise разрешается с объектом XMLHttpRequest,
 * содержащим загруженные данные в свойстве responseText.
 *
 * Функция поддерживает загрузку как с локальных путей (file://),
 * так и с удаленных URL (http://, https://).
 *
 * @param {string} url - URL-адрес или путь к файлу для загрузки данных
 * @returns {Promise<XMLHttpRequest>} Promise, который разрешается объектом XMLHttpRequest при успешной загрузке
 *
 * @throws {Error} Если URL не указан ('Не указан путь к файлу!')
 * @throws {Error} Если произошел таймаут запроса ('The request for [url] timed out')
 * @throws {Error} Если произошла сетевая ошибка ('Network error occurred while fetching [url]')
 * @throws {Error} Если сервер вернул ошибку ('Ошибка HTTP: [status] [statusText]')
 *
 * @example
 * // Загрузка данных и обработка результата
 * loadData('https://example.com/data.csv')
 *   .then(xhr => {
 *     processData(xhr.responseText);
 *   })
 *   .catch(error => {
 *     console.error('Ошибка загрузки:', error);
 *   });
 *
 * @example
 * // Использование с async/await
 * async function loadAndProcessData() {
 *   try {
 *     const xhr = await loadData('https://example.com/data.csv');
 *     processData(xhr.responseText);
 *   } catch (error) {
 *     console.error('Ошибка загрузки:', error);
 *   }
 * }
 */
function loadData(url) {
  // Проверка наличия URL. Если URL не передан или пустой, выбрасываем исключение
  if (!url) {
    throw new Error('Не указан путь к файлу!');
  }

  // Возвращаем Promise для асинхронной загрузки данных
  return new Promise((resolve, reject) => {
    // Создаем объект XMLHttpRequest для выполнения HTTP-запроса
    const xhr = new XMLHttpRequest();

    /**
     * Обработчик таймаута запроса - вызывается, если запрос не завершился за отведенное время
     * @private
     */
    xhr.ontimeout = function () {
      reject(new Error('The request for ' + url + ' timed out.'));
    };

    /**
     * Обработчик ошибки запроса - вызывается при сетевых ошибках
     * @private
     */
    xhr.onerror = function () {
      reject(new Error('Network error occurred while fetching ' + url));
    };

    /**
     * Обработчик успешного завершения запроса - проверяет код HTTP-статуса
     * и разрешает или отклоняет Promise соответственно
     * @private
     */
    xhr.onload = function () {
      // Проверяем, что запрос завершен (readyState === 4 означает DONE)
      if (xhr.readyState === 4) {
        if (xhr.status === 200) {
          // Код 200 означает успешный запрос
          resolve(xhr);
        } else {
          // Любой другой код - ошибка HTTP
          reject(new Error('Ошибка HTTP: ' + xhr.status + ' ' + xhr.statusText));
        }
      }
    };

    // Настройка и отправка запроса
    xhr.open('GET', url, true); // true означает асинхронный запрос
    xhr.timeout = 60 * 1000; // Таймаут 60 секунд (в миллисекундах)
    xhr.send(null); // null т.к. это GET-запрос без данных
  });
}

/**
 * Преобразует строку CSV в массив объектов JavaScript
 *
 * @description
 * Функция парсит входную строку CSV и преобразует её в массив объектов JavaScript.
 * Первая строка CSV рассматривается как заголовки столбцов, которые станут свойствами объектов.
 * Остальные строки преобразуются в объекты, где ключами являются заголовки столбцов.
 *
 * Функция обрабатывает различные форматы окончания строк (\r\n, \n, \r) и
 * автоматически удаляет пробельные символы из заголовков столбцов.
 *
 * Если количество значений в строке меньше количества заголовков,
 * недостающие значения будут установлены как undefined.
 * Если количество значений больше заголовков, лишние значения будут игнорироваться.
 *
 * @param {string} csvString - Строка в формате CSV для парсинга
 * @param {string} [delimiter=';'] - Символ-разделитель полей в CSV (по умолчанию ';')
 * @returns {Array<Object>} - Массив объектов, представляющих строки CSV
 * @throws {Error} - Если входная строка пуста или не содержит строк данных
 *
 * @example
 * const csvData = 'name;age\nJohn;30\nJane;25';
 * const result = parseCsv(csvData);
 * // Результат: [{ name: 'John', age: '30' }, { name: 'Jane', age: '25' }]
 *
 * @example
 * // С другим разделителем
 * const csvData = 'name,age\nJohn,30\nJane,25';
 * const result = parseCsv(csvData, ',');
 */
function parseCsv(csvString, delimiter = ';') {
  // Проверяем, что входная строка не пуста
  if (!csvString || typeof csvString !== 'string') {
    throw new Error('CSV строка не предоставлена или имеет неверный формат');
  }

  // Разбиваем строку на отдельные строки, учитывая возможные разные окончания строк
  // Регулярное выражение /[\r\n]+/g обрабатывает \r\n, \n и \r как разделители строк
  const lines = csvString.split(/[\r\n]+/g);

  // Проверяем, что есть хотя бы заголовок и одна строка данных
  if (lines.length < 2) {
    throw new Error('CSV строка должна содержать как минимум заголовок и одну строку данных');
  }

  // Получаем заголовки из первой строки и очищаем их от пробелов
  const headers = lines[0].split(delimiter).map((header) => header.trim());

  // Преобразуем остальные строки в объекты
  const data = lines
    .slice(1)
    .map((line) => {
      // Пропускаем пустые строки
      if (!line.trim()) {
        return null;
      }

      // Разделяем строку по указанному разделителю
      const values = line.split(delimiter);

      // Создаем объект, где ключи - это заголовки, а значения - соответствующие значения из текущей строки
      // Используем reduce для преобразования массивов заголовков и значений в единый объект
      return headers.reduce((obj, header, index) => {
        // Присваиваем значение из текущей строки соответствующему свойству объекта
        // Если значение не существует (index >= values.length), будет использовано undefined
        obj[header] = values[index];
        return obj;
      }, {});
    })
    .filter(Boolean); // Удаляем пустые строки (null)

  return data;
}

/*******************************************************************
 *
 *                            Утилиты
 *
 ********************************************************************/

/**
 * Отображает диалоговое окно с сообщением
 *
 * @description
 * Функция проверяет доступность объекта Common.UI и, если он существует,
 * показывает диалоговое окно с указанным заголовком и сообщением.
 * Используется для вывода информации пользователю в R7 Office.
 *
 * @param {string} title - Заголовок диалогового окна
 * @param {string} msg - Текст сообщения в диалоговом окне
 * @returns {void}
 *
 * @example
 * // Показать информационное сообщение
 * uiAlert('Информация', 'Операция выполнена успешно');
 *
 * @example
 * // Показать сообщение об ошибке
 * uiAlert('Ошибка', 'Не удалось загрузить данные');
 */
function uiAlert(title, msg) {
  // Проверяем доступность интерфейсного объекта Common.UI
  if (typeof Common !== 'undefined' && Common.UI) {
    // Вызываем метод alert для отображения диалогового окна
    Common.UI.alert({
      title, // Заголовок окна
      msg, // Текст сообщения
    });
  }
  // Если Common.UI недоступен, ничего не делаем
}

/**
 * Преобразует путь Windows в URL-формат для загрузки файлов
 *
 * @description
 * Функция конвертирует путь Windows в формат URL file:///, который можно использовать
 * для загрузки локальных файлов через XMLHttpRequest. Обрабатывает различные
 * форматы Windows-путей, включая длинные имена файлов и различные разделители.
 *
 * @param {string} windowsPath - Путь в формате Windows (с backslash-разделителями)
 * @returns {string} URL в формате file:/// для доступа к локальному файлу
 *
 * @example
 * // Преобразование обычного пути Windows
 * const url = convertPath('C:\\data\\file.csv');
 * // Результат: file:///C:/data/file.csv
 *
 * @example
 * // Обработка длинного пути Windows
 * const url = convertPath('\\\\?\\C:\\Very Long Path\\file.csv');
 * // Результат: file:///C:/Very Long Path/file.csv
 */
function convertPath(windowsPath) {
  // Обрабатываем специальный случай длинных имен файлов Windows
  // Префикс \\?\ используется для длинных путей (более 260 символов)
  // См.: https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file#short-vs-long-names
  windowsPath = windowsPath.replace(/^\\\\\?\\/, '');

  // Конвертируем обратные слеши (\) в прямые (/) для совместимости с URL
  // В именах файлов Windows не могут одновременно использоваться оба разделителя
  windowsPath = windowsPath.replace(/\\/g, '/');

  // Сжимаем последовательности слешей (//, ///) в один слеш (/)
  // Это безопасная операция в POSIX-путях и предотвращает ошибки при объединении путей
  windowsPath = windowsPath.replace(/\/\/+/g, '/');

  // Удаляем префикс file:///, если он уже присутствует, чтобы избежать дублирования
  windowsPath = windowsPath.replace(/^file:\/+/, '');

  // Возвращаем путь с префиксом file:/// для использования в XMLHttpRequest
  return `file:///${windowsPath}`;
}

/**
 * Преобразует серийный номер даты из формата R7 Office в JavaScript Date
 *
 * @description
 * В R7 Office даты представлены в виде последовательных серийных номеров,
 * где 1 соответствует 1 января 1900 года. Эта функция конвертирует такой
 * серийный номер в объект JavaScript Date для дальнейших манипуляций с датой.
 *
 * @param {number} r7SerialDate - Серийный номер даты в формате R7 Office
 * @returns {Date} Объект JavaScript Date, соответствующий входному серийному номеру
 *
 * @example
 * // Преобразование даты (1 января 2020 года примерно соответствует серийному номеру 43831)
 * const jsDate = r7SerialToJsDate(43831);
 * console.log(jsDate.toISOString()); // 2020-01-01T00:00:00.000Z
 *
 * @example
 * // Работа с пустыми значениями
 * const defaultDate = r7SerialToJsDate(null) || new Date(2000, 0, 1);
 */
function r7SerialToJsDate(r7SerialDate) {
  // В формате даты R7 Office 1 соответствует 1 января 1900 года
  // Вычитаем 1, чтобы правильно преобразовать день в UTC
  return new Date(Date.UTC(0, 0, r7SerialDate - 1));
}

/**
 * Создает диапазон ячеек по начальной позиции и размеру
 *
 * @description
 * Функция упрощает создание прямоугольного диапазона ячеек в электронной таблице.
 * Вместо определения верхней левой и нижней правой ячеек напрямую, позволяет
 * указать начальную позицию и размеры диапазона (количество строк и столбцов).
 *
 * @param {Object} sheet - Объект листа, в котором создается диапазон
 * @param {number} row - Индекс начальной строки (нумерация с нуля)
 * @param {number} col - Индекс начального столбца (нумерация с нуля)
 * @param {number} [rows=1] - Количество строк в диапазоне (по умолчанию 1)
 * @param {number} [cols=1] - Количество столбцов в диапазоне (по умолчанию 1)
 * @returns {Object} Объект диапазона ячеек
 *
 * @example
 * // Создать диапазон из 3 строк и 4 столбцов, начиная с ячейки B2 (строка 1, столбец 1)
 * const range = getRangeBySize(sheet, 1, 1, 3, 4);
 *
 * @example
 * // Создать диапазон размером 1x2 (одна строка, два столбца)
 * const headerRange = getRangeBySize(sheet, 0, 0, 1, 2);
 * headerRange.SetValue(['Имя', 'Возраст']);
 */
function getRangeBySize(sheet, row, col, rows = 1, cols = 1) {
  // Определяем верхнюю левую ячейку диапазона
  const leftTop = sheet.GetRangeByNumber(row, col);

  // Определяем нижнюю правую ячейку диапазона
  // Вычитаем 1, т.к. размер включает начальную ячейку
  const rightBottom = sheet.GetRangeByNumber(row + rows - 1, col + cols - 1);

  // Создаем и возвращаем диапазон между этими двумя ячейками
  return sheet.GetRange(leftTop, rightBottom);
}
