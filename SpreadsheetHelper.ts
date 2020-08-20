namespace SpreadsheetHelper {
  const ss = SpreadsheetApp.getActive()

  export const LineSymbol = Symbol('line')
  export const IndexSymbol = Symbol('index')

  export type CellVal = string | number | boolean | Date

  export interface RowObj {
    [key: string]: CellVal | null
  }

  // TODO Как указать тип кастомного трансформера?
  // https://stackoverflow.com/questions/56505560/could-be-instantiated-with-a-different-subtype-of-constraint-object
  // export type SheetRowsTransformer<T> = (
  //   row: CellVal[],
  //   index: number,
  //   sheet: {
  //     header: string[]
  //     rows: CellVal[][]
  //   }
  // ) => T

  export interface SheetRowObj extends RowObj {
    [LineSymbol]: number
    [IndexSymbol]: number
  }

  export const defaultSheetRowsTransformer = (
    row: CellVal[],
    index: number,
    sheet: {
      header: string[]
    }
  ) => {
    const rowObj = row.reduce((res, val, colIndex) => {
      res[sheet.header[colIndex]] = val === '' ? null : val
      res[IndexSymbol] = index
      res[LineSymbol] = index + 1 // TODO Use frozen rows
      return res
    }, {} as SheetRowObj)

    return rowObj
  }

  export function getSheetTableRange(sheetName: string) {
    const sheet = ss.getSheetByName(sheetName)

    if (!sheet) {
      throw new Error(`Лист "${sheetName}" не найден.`)
    }

    const frozenRows = sheet.getFrozenRows()
    const lastRow = sheet.getLastRow()
    const lastCol = sheet.getLastColumn()

    if (frozenRows === 0) {
      throw new Error(`Заголовок листа "${sheetName}" должен быть закреплен.`)
    } else if (frozenRows > 1) {
      throw new Error(
        `Заголовок листа "${sheetName}" должен быть одной строкой.`
      )
    }

    if (lastRow === 0 || lastCol === 0) {
      throw new Error(`Лист "${sheetName}" не заполнен.`)
    }

    let range
    try {
      range = sheet.getRange(1, 1, lastRow, lastCol)
    } catch (err) {
      throw new Error(
        `Ошибка получения диапазона .getRange(1, 1, ${lastRow}, ${lastCol})` +
          ` для листа "${sheetName}"`
      )
    }

    return range
  }

  export function getSheetRowsAsObjects(
    range: GoogleAppsScript.Spreadsheet.Range
  ): SheetRowObj[]

  export function getSheetRowsAsObjects(sheetName: string): SheetRowObj[]

  export function getSheetRowsAsObjects(
    sheet: string | GoogleAppsScript.Spreadsheet.Range
    // transformer = defaultSheetRowsTransformer
  ): SheetRowObj[] {
    const transformer = defaultSheetRowsTransformer // TODO custom

    const range = typeof sheet === 'string' ? getSheetTableRange(sheet) : sheet
    const sheetName =
      typeof sheet === 'string' ? sheet : sheet.getSheet().getName()

    const values = range.getValues() as CellVal[][]

    if (values.length === 1) {
      return []
    }

    const header = values[0] as string[]

    //#region Header test
    const uniqHeaders = Array.from(new Set(header))

    if (header.length !== uniqHeaders.length) {
      throw new Error(
        `Заголовки в таблице "${sheetName}" не должны повторяться`
      )
    }

    if (uniqHeaders.some(h => typeof h !== 'string')) {
      throw new Error(`Заголовки в таблице "${sheetName}" должны быть строками`)
    }

    if (uniqHeaders.some(h => h === '')) {
      throw new Error(
        `Заголовки в таблице "${sheetName}" должны быть заполнены`
      )
    }
    //#endregion

    const rowsAsObjects: ReturnType<typeof transformer>[] = []
    for (let i = 1; i < values.length; i++) {
      rowsAsObjects.push(
        transformer(values[i], i, {
          header
        })
      )
    }

    return rowsAsObjects
  }

  export interface SheetObjectParams {
    keyHeaders?: string[]
    rowsFilter?: (row: SheetRowObj) => boolean
  }

  class RowObject {
    constructor(private row: SheetRowObj, private sheet: SheetObject) {}

    getLine() {
      return this.row[LineSymbol]
    }

    getIndex() {
      return this.row[IndexSymbol]
    }

    getRow() {
      return this.row
    }

    getRange(headerName: string) {
      return this.sheet.range.getCell(
        this.row[LineSymbol],
        this.sheet.getHeaderCol(headerName)
      )
    }

    getValue(headerName: string) {
      return this.getRange(headerName).getValue()
    }

    setValue(headerName: string, value: CellVal) {
      this.getRange(headerName).setValue(value == null ? '' : value)
      return this
    }
  }

  export class SheetObject {
    /** Ключевые заголовки листа */
    private keyHeaders: string[] | undefined

    /** Строки по ключам */
    private rowsByKeysMap = new Map<string, RowObject>()

    /** Заголовок */
    private header: string[]

    /** Колонка заголовка по его наименованию */
    private headerColMap = new Map<string, number>()

    /** Последняя строка */
    private lastRow: number

    /** Наименование листа таблицы */
    sheetName: string

    /** Диапазон таблицы листа */
    range: GoogleAppsScript.Spreadsheet.Range

    /** Объекты строк листа */
    rows: RowObject[] // ReturnType<typeof getSheetRowsAsObjects>

    /** Возвращает ключ строки для индекса */
    private getRowKey(row: RowObject) {
      return JSON.stringify(
        this.keyHeaders?.reduce(
          (res, h) => res.concat(row.getRow()[h]),
          [] as Array<CellVal | null>
        )
      )
    }

    private addToIndex(row: RowObject) {
      const rowKey = this.getRowKey(row)

      if (this.rowsByKeysMap.has(rowKey)) {
        const existRow = this.rowsByKeysMap.get(rowKey)!

        throw new Error(
          `В таблице "${this.sheetName}"` +
            (this.keyHeaders!.length > 1
              ? ` по ключевым полям ${JSON.stringify(this.keyHeaders)}`
              : ` по ключевому полю "${this.keyHeaders![0]}"`) +
            ` строка ${existRow.getLine()} аналогична строке ${row.getLine()}`
        )
      } else {
        this.rowsByKeysMap.set(rowKey, row)
      }
    }

    constructor(sheetName: string, params: SheetObjectParams = {}) {
      this.sheetName = sheetName

      this.range = getSheetTableRange(sheetName)

      let rowsFilter = params.rowsFilter ?? (() => true)

      this.rows = getSheetRowsAsObjects(this.range)
        .filter(it => rowsFilter(it))
        .map(row => new RowObject(row, this))

      // TODO Таблица может быть заполнена пустыми строками
      this.lastRow = this.range.getLastRow()

      this.header = this.range
        .offset(0, 0, 1, this.range.getNumColumns())
        .getValues()[0]

      this.header.forEach((h, index) => {
        this.headerColMap.set(h, index)
      })

      // Индекс
      if (params.keyHeaders?.length) {
        this.keyHeaders = params.keyHeaders

        this.rows.forEach(row => {
          this.addToIndex(row)
        })
      }
    }

    /** Является ли таблица индексированной */
    indexed() {
      return !!this.keyHeaders?.length
    }

    /** Возвращает колонку для указанного заголовка */
    getHeaderCol(headerName: string) {
      if (this.headerColMap.has(headerName)) {
        return this.headerColMap.get(headerName)! + 1
      } else {
        throw new Error(`Заголовок "${headerName}" не найден`)
      }
    }

    getRowByKey(key: CellVal | CellVal[]) {
      if (!this.indexed()) {
        throw new Error(
          `Нельзя получить строку по ключу -` +
            ` таблица создана без указания индекса`
        )
      }

      const rowKeys = key instanceof Array ? key : [key]

      if (rowKeys.length !== this.keyHeaders?.length) {
        throw new Error(
          `Нельзя получить строку по ключу из ${rowKeys.length} заголовков` +
            ` таблица создана с индексом длинной ${this.keyHeaders?.length} заголовков`
        )
      }

      return this.rowsByKeysMap.get(JSON.stringify(rowKeys))
    }

    /** Добавить строки в таблицу */
    addRows(rows: RowObj[]) {
      const values = rows.map(row => {
        return this.header.map(h => {
          const cellVal = row[h]
          return cellVal == null ? '' : cellVal
        })
      })

      const startRow = this.lastRow

      // Сохраняем в таблицу до валидации на индекс.
      // Преимущества:
      //  - проще вывести ошибку по номеру строки
      //  - не теряются сохраняемые данные (важно?)
      //  - проще отладить
      // Недостатки:
      //  - в таблицу записываются не уникальные данные
      this.range
        .offset(startRow, 0, values.length, values[0].length)
        .setValues(values)

      this.lastRow += values.length

      // TODO Доработать. Пропускать через трансформер.
      const rowObjects = rows.map(
        (row, index) =>
          new RowObject(
            {
              [IndexSymbol]: startRow + index, // TODO Нет ошибки с индексом?
              [LineSymbol]: startRow + index + 1,
              ...row
            },
            this
          )
      )

      this.rows = this.rows.concat(rowObjects)

      // TODO Поддержать формулы (указывать/протягивать возможно?)
      if (this.indexed()) {
        rowObjects.forEach(rowObject => {
          this.addToIndex(rowObject)
        })
      }

      return rowObjects
    }
  }
}
