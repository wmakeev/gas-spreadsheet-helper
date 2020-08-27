namespace SpreadsheetHelper {
  export type CellValue = string | number | boolean | Date
  export type RowRecordValue = CellValue | null

  export type RowRecord = Record<string, RowRecordValue>

  export interface SheetTableOptions<T extends RowRecord> {
    /** Spreadsheet id if not active table */
    spreadsheetId?: string

    /** Ensure table headers */
    mandatoryHeaders?: Set<keyof T>

    /** Headers in uniq row key */
    keyHeaders?: Array<keyof T>
  }

  export class SheetTable<T extends RowRecord> {
    public spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
    public sheet: GoogleAppsScript.Spreadsheet.Sheet

    private lastRow: number
    // private lastColumn: number
    private range: GoogleAppsScript.Spreadsheet.Range | null = null

    /** Headers column numbers */
    private headerColumn = new Map<keyof T, number>()
    private headers: Array<keyof T>

    private rowByKeyMap = new Map<string, SheetRow<T>>()
    private rowByRecordWeakMap = new WeakMap<T, SheetRow<T>>()

    private rows = new Set<SheetRow<T>>()

    constructor(
      public sheetName: string,
      public options: SheetTableOptions<T> = {}
    ) {
      // Get spreadsheet
      if (options.spreadsheetId) {
        this.spreadsheet = SpreadsheetApp.openById(options.spreadsheetId)
      } else {
        this.spreadsheet = SpreadsheetApp.getActive()
      }

      // Get sheet
      const sheet = this.spreadsheet.getSheetByName(sheetName)

      if (sheet) {
        this.sheet = sheet
      } else {
        throw new Error(`Sheet "${sheetName}" not exist"`)
      }

      // Header
      const frozenRows = this.sheet.getFrozenRows()

      if (frozenRows === 0) {
        throw new Error(`Заголовок листа "${sheetName}" должен быть закреплен.`)
      } else if (frozenRows > 1) {
        throw new Error(
          `Заголовок листа "${sheetName}" должен быть одной строкой.`
        )
      }

      const range = this.getSheetRange(true)

      const lastRow = (this.lastRow = range.getLastRow())
      const lastColumn = range.getLastColumn()

      if (lastRow === 0 || lastColumn === 0) {
        throw new Error(`Лист "${sheetName}" не заполнен.`)
      }

      const sheetValues = range.getValues() as CellValue[][]

      const headers = sheetValues.shift() as Array<keyof T>

      //#region Header test
      if (headers.some(h => h === '')) {
        throw new Error(
          `Заголовки в таблице "${this.sheetName}" должны быть заполнены`
        )
      }

      if (headers.some(h => typeof h !== 'string')) {
        throw new Error(
          `Заголовки в таблице "${this.sheetName}" должны быть строками`
        )
      }

      const uniqHeaders = Array.from(new Set(headers))

      if (headers.length !== uniqHeaders.length) {
        throw new Error(
          `Заголовки в таблице "${this.sheetName}" не должны повторяться`
        )
      }
      //#endregion

      this.headers = headers

      this.headerColumn = headers.reduce((res, h, index) => {
        res.set(h, index + 1)
        return res
      }, new Map<keyof T, number>())

      // Check for mandatory headers
      if (options.mandatoryHeaders?.size || options.keyHeaders?.length) {
        const keyHeaders = options.keyHeaders ?? []
        const mandatoryHeaders = Array.from(options.mandatoryHeaders ?? [])

        const allMandatoryHeaders = Array.from(
          new Set(keyHeaders.concat(mandatoryHeaders))
        )

        const lostHeaders = allMandatoryHeaders.filter(
          h => this.headerColumn.get(h) == null
        )

        if (lostHeaders.length) {
          throw (
            new Error(
              `В таблице "${this.sheetName}" отсутствуют необходимые заголовки -` +
                lostHeaders.map(h => `"${h}"`).join(', ')
            ) + '.'
          )
        }
      }

      //#region Загрузка строк таблицы
      let sheetRowsRecords: T[] = []
      for (let sheetRowValues of sheetValues) {
        let rowRecord = {} as T
        for (let key of this.headerColumn.keys()) {
          const val: any =
            sheetRowValues[this.headerColumn.get(key)! - 1] ?? null
          rowRecord[key] = val === '' ? null : val
        }
        sheetRowsRecords.push(rowRecord)
      }

      this.createInternalRows(sheetRowsRecords, 1)
      //#endregion
    }

    /** Return row record index hash */
    private getRecordHash(record: T | RowRecordValue[]) {
      if (!this.options.keyHeaders?.length) {
        throw new Error('Ключевые заголовки в таблице не заданы')
      }

      const keyValues: RowRecordValue[] =
        record instanceof Array
          ? record
          : this.options.keyHeaders.reduce((res, h) => {
              res.push(record[h])
              return res
            }, [] as RowRecordValue[])

      return JSON.stringify(keyValues) // TODO Hash function
    }

    private createInternalRows(records: T[], lastRow: number): SheetRow<T>[] {
      const rows = records.map(
        (it, index) =>
          new SheetRow(this, it, lastRow + index + 1, row =>
            this.attachRow(row)
          )
      )

      return rows
    }

    // TODO Думаю строка может быть и не аттачнута. Без id (номера строки)..
    // затем можно аттачить уже отдельно

    /** Добавить строку во внутреннее представление таблицы */
    private attachRow(row: SheetRow<T>) {
      const record = row.getRecord()

      this.rowByRecordWeakMap.set(record, row)

      let isNewRow = !this.rows.has(row)

      if (isNewRow) this.rows.add(row)

      if (this.options.keyHeaders?.length) {
        // Строка не добавляется в индекс если все ключевые заголовки пусты
        if (this.options.keyHeaders.every(h => record[h] == null)) {
          return
        }

        const recHash = this.getRecordHash(record)

        if (isNewRow && this.rowByKeyMap.has(recHash)) {
          const existRow = this.rowByKeyMap.get(recHash)!

          throw new Error(
            `В таблице "${this.sheetName}"` +
              (this.options.keyHeaders.length > 1
                ? ` по ключевым полям ${JSON.stringify(
                    this.options.keyHeaders
                  )}`
                : ` по ключевому полю "${this.options.keyHeaders![0]}"`) +
              ` строка ${existRow.getLine()} аналогична строке ${row.getLine()}`
          )
        } else {
          this.rowByKeyMap.set(recHash, row)
        }
      }
    }

    private getValues(record: T) {
      return this.headers.map(h => {
        const cellVal = record[h]
        return cellVal == null ? '' : cellVal
      })
    }

    /** Returns sheet range */
    getSheetRange(forceUpdate = false) {
      if (!forceUpdate && this.range) {
        return this.range
      }

      let range

      const lastRow = this.sheet.getLastRow()
      const lastColumn = this.sheet.getLastColumn()

      try {
        range = this.sheet.getRange(
          1,
          1,
          this.sheet.getLastRow(),
          this.sheet.getLastColumn()
        )
      } catch (err) {
        throw new Error(
          `Ошибка получения диапазона .getRange(1, 1, ${lastRow}, ${lastColumn})` +
            ` для листа "${this.sheetName}"`
        )
      }

      this.range = range

      return range
    }

    /** Is table has index */
    indexed() {
      return !!this.options.keyHeaders?.length
    }

    /** Возвращает колонку для указанного заголовка */
    getHeaderColumn<P extends keyof T>(headerName: P) {
      if (this.headerColumn.has(headerName)) {
        return this.headerColumn.get(headerName)
      } else {
        throw new Error(`Заголовок "${headerName}" не найден`)
      }
    }

    /** Returns corresponding record row */
    getRowByRecord(record: T) {
      return this.rowByRecordWeakMap.get(record)
    }

    getRowByKey(key: RowRecordValue | RowRecordValue[]) {
      if (!this.indexed()) {
        throw new Error(
          `Нельзя получить строку по ключу -` +
            ` таблица создана без указания индекса`
        )
      }

      const rowKeys = key instanceof Array ? key : [key]

      if (rowKeys.length !== this.options.keyHeaders?.length) {
        throw new Error(
          `Нельзя получить строку по ключу из ${rowKeys.length} заголовков` +
            ` таблица создана с индексом длинной ${this.options.keyHeaders?.length} заголовков`
        )
      }

      const rowKey = this.getRecordHash(rowKeys)

      return this.rowByKeyMap.get(rowKey)
    }

    getRows() {
      return Array.from(this.rows.values())
    }

    /** Добавить строки в таблицу */
    addRecords(records: T[]) {
      // Преобразование массива объектов в строки таблицы
      const values = records.map(record => this.getValues(record))

      const range = this.getSheetRange()

      const lastRow = range.getLastRow()

      // Сохраняем в таблицу до валидации на индекс.
      // Преимущества:
      //  - проще вывести ошибку по номеру строки
      //  - не теряются сохраняемые данные (важно?)
      //  - проще отладить
      // Недостатки:
      //  - в таблицу записываются не уникальные данные
      range
        .offset(this.lastRow, 0, values.length, values[0].length)
        .setValues(values)

      this.lastRow = this.getSheetRange(true).getLastRow()

      const newRows = this.createInternalRows(records, lastRow)

      return newRows
    }
  }

  export class SheetRow<T extends RowRecord> {
    constructor(
      public sheetTable: SheetTable<T>, // GoogleAppsScript.Spreadsheet.Sheet,
      private rowRecord: T,
      private rowLine: number,
      private changeHook: (row: SheetRow<T>) => void
    ) {
      this.changeHook(this)
    }

    getRecord() {
      return this.rowRecord
    }

    getCell<P extends keyof T>(headerName: P) {
      return this.sheetTable
        .getSheetRange()
        .getCell(this.getLine(), this.sheetTable.getHeaderColumn(headerName)!)
    }

    getValue<P extends keyof T>(header: P): T[P] {
      return this.rowRecord[header]
    }

    setValue<P extends keyof T>(header: P, value: T[P]) {
      const cell = this.getCell(header)

      cell.setValue(value == null ? '' : value)

      this.rowRecord = {
        ...this.rowRecord,
        [header]: value
      }

      this.changeHook(this)

      return this
    }

    getLine() {
      return this.rowLine
    }
  }
}
