class Attendance {
  /**
   * @constructor 出席情報のシートを操作します
   * @param {SpreadsheetApp.Spreadsheet} spreadSheet - 対象となるスプレッドシート
   * @param {string} name - シート名
   */
  constructor(spreadSheet, name) {
    this.spreadSheet = spreadSheet;
    this.sheet = spreadSheet.getSheetByName(name);
    this.name = name;
  }

  /**
   * シートを取得します
   * @return {SpreadsheetApp.Sheet} シート
   */
  getSheet() {
    return this.sheet;
  }

  /**
   * シートが存在するかを確認します
   * @return {boolean} 存在するか
   */
  exists() {
    return Boolean(this.sheet);
  }

  /**
   * シートを初期化(新規作成)します
   */
  initialize() {
    // シート作成
    const sheet = this.spreadSheet.insertSheet(this.name);

    // 「名前」を入力
    const name = sheet.getRange(1, 1);
    name.setValue("名前");
    name.setFontSize(14);
    name.setBorder(true, true, true, false, false, false, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // 「割合」を入力
    const attendanceRate = sheet.getRange(1, 2);
    attendanceRate.setValue("割合");
    attendanceRate.setFontSize(14);
    attendanceRate.setBorder(true, false, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    attendanceRate.setBorder(null, true, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
    
    // 「全体」を入力
    const whole = sheet.getRange(3, 1);
    whole.setValue("全体");
    whole.setFontSize(14);
    sheet.setRowHeight(3, 35);

    // 列の枠線を設定
    const rows = sheet.getRange(3, 1, 1, 2);
    rows.setBorder(true, true, true, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    rows.setBorder(null, null, null, null, true, true, null, SpreadsheetApp.BorderStyle.SOLID);

    // 固定行・列設定
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(2);
    sheet.setColumnWidths(3, sheet.getMaxColumns() - 2, 32);
    sheet.setColumnWidth(2, 45);
    sheet.getRange(2, 3, 2, 50).setHorizontalAlignment('center');
    sheet.getRange("A1:A2").merge();
    sheet.getRange("B1:B2").merge();

    // 条件付き書式を設定
    const rate = sheet.getRange("B3:ZZZ1000");
    const lessThan80 = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(80)
      .setBackground("#b7e1cd")
      .setRanges([rate])
      .build();

    // 条件付き書式を設定
    const attendance = sheet.getRange("C3:ZZZ1000")
    const absentCondition = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("欠席")
      .setBackground("#f4c7c3")
      .setRanges([attendance])
      .build();
    const lateCondition = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("遅刻")
      .setBackground("#fce8b2")
      .setRanges([attendance])
      .build();
    const earlyLeavingCondition = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("早退")
      .setBackground("#fce8b2")
      .setRanges([attendance])
      .build();

    // 適応
    const rules = sheet.getConditionalFormatRules();
    rules.push(lessThan80, absentCondition, lateCondition, earlyLeavingCondition);
    sheet.setConditionalFormatRules(rules);

    // パートを挿入
    for (let i in Part) {
      const lastRow = sheet.getLastRow();
      sheet.insertRowBefore(lastRow);
      sheet.setRowHeight(lastRow, 50)

      const partRange = sheet.getRange(lastRow, 1);
      partRange.setValue(Part[i]);
    }

    this.sheet = sheet;
  }

  /**
   * 月のセル取得します
   * @param {Number} month - 月
   * @return {SpreadsheetApp.Range|null} 月のセル(ない場合はnullを返却)
   */
  getMonth(month) {
    // 最終列取得
    const lastColumn = this.sheet.getLastColumn();

    // 最終列が2以下(＝月が無い)の場合nullを返却
    if (lastColumn <= 2) {
      return null;
    }

    // 月の行のセルの値を取得
    const monthValues = this.sheet.getRange(1, 3, 1, lastColumn - 2).getValues()[0];

    // ループ
    for (let i in monthValues) {

      // 値が空だったらスキップ
      if (monthValues[i] == "") {
        continue;
      }

      // 値が指定された月だったら
      if (monthValues[i] == month) {

        // 月の1つめのセルを取得
        const monthRange = this.sheet.getRange(1, Number(i) + 3);

        // もしセル結合があったら
        if (monthRange.isPartOfMerge()) {

          // 結合を取得し、その範囲のセルを返却
          return monthRange.getMergedRanges()[0];
        }

        // なければそのまま返却
        return monthRange;
      }
    }

    // 月が見つからなければnullを返却
    return null;
  }

  /**
   * 月のセルをすべて取得します
   * @return {List.<Month>} Month型のリスト
   */
  getAllMonth() {
    // オブジェクトを初期化
    let monthRangeList = [];

    // 1から12月までループ
    for (let i = 4; i <= 15; i++) {

      let j = i;
      if (i == 13) {
        j = 1;
      } else if (i == 14) {
        j = 2;
      } else if (i == 15) {
        j = 3;
      }

      // 月のセルを取得
      const monthRange = this.getMonth(j);

      // 取得内容がnullでなければmonthRangeListに追加
      if (monthRange != null) {
        monthRangeList.push(new MonthRange(j, monthRange));
      }
    }

    // monthRangeListを返却
    return monthRangeList;
  }

  /**
   * 日のセル取得します
   * @param {Date} ddate - Date型
   * @return {SpreadsheetApp.Range|null} 月のセル(ない場合はnullを返却)
   */
  getDate(ddate) {
    const month = ddate.getMonth() + 1;
    const date = ddate.getDate();

    // 月のセルを取得
    const monthRange = this.getMonth(month);

    // もし取得内容が空だったらnullを返却
    if (!monthRange) {
      return null;
    }

    // 月のセルから、日のセルの範囲の値を取得
    const dayValues = this.sheet.getRange(2, monthRange.getColumn(), 1, monthRange.getNumColumns()).getValues()[0];

    // 日をループ
    for (let i in dayValues) {

      // もし値が空だったらスキップ
      if (dayValues[i] == "") {
        continue;
      }

      // もし値が指定された日だったら
      if (dayValues[i] == date) {

        // その日のセルを返却
        return this.sheet.getRange(2, Number(i) + monthRange.getColumn());
      }
    }

    // 日が見つからなければnullを返却
    return null;
  }

  /**
   * 月のセルを挿入します(既に存在する場合は使用しないでください)
   * @param {Number} month - 月
   * @return {SpreadsheetApp.Range|null} 月のセル
   */
  insertMonth(month) {
    // 最終列を取得
    const lastColumn = this.sheet.getLastColumn();

    // 1月2月3月を後ろに持っていくため、処理用に13/14/15にする
    let conductMonth = month;
    if (conductMonth == 1) {
      conductMonth = 13;
    } else if (conductMonth == 2) {
      conductMonth = 14;
    } else if (conductMonth == 3) {
      conductMonth = 15;
    }

    /** 月が順番になるように追加する処理 */
    // 最終列が2以下(＝月が無い)でない場合
    if (lastColumn > 2) {
      // すべての月のセルを取得
      const monthRanges = this.getAllMonth();

      // 列数を計算するために初期化
      let totalColumn = 0;

      // 月をループ
      for (let i in monthRanges) {

        // その月の値を取得
        const value = monthRanges[i].range.getValue();

        // 値が空だったらスキップ
        if (value == "") {
          continue;
        }

        // 列数に追加
        totalColumn += monthRanges[i].range.getNumColumns();
        console.log(monthRanges[i].range.getNumColumns())

        // もし処理用の月が値より小さかったら
        if (value < conductMonth) {

          // 後ろに月がある場合
          if (monthRanges[Number(i) + 1]) {

            // 後ろの月の値が処理用の月より小さかったらスキップ
            if (monthRanges[Number(i) + 1].range.getValue() < conductMonth) {
              continue;
            }
          }

          /** それ以外の場合 */
          // 月のカラムの後ろに列を追加
          this.sheet.insertColumnAfter(totalColumn + 2);

          // 追加した範囲を指定
          const monthRange = this.sheet.getRange(1, totalColumn + 3);

          // 範囲に値・フォントサイズ・配置・枠線を設定
          monthRange.setValue(month);
          monthRange.setFontSize(14);
          monthRange.setHorizontalAlignment("left");
          monthRange.setBorder(true, true, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
          monthRange.setBorder(null, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);

          // 割合をいれる範囲を指定
          const dateRange = this.sheet.getRange(2, monthRange.getColumn());

          // 割合の範囲に値・フォントサイズ・配置・枠線を設定
          dateRange.setValue("割合");
          dateRange.setFontSize(10);
          dateRange.setHorizontalAlignment("center");
          dateRange.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
          dateRange.setBorder(null, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
          dateRange.setBackground(null);

          // 実際の割合の入るセルを指定
          const rate = this.sheet.getRange(3, monthRange.getColumn(), 1000);
          // その月のセルを返却
          return monthRange;
        }
      }
    }

    /** それ以外の場合(=その月が既存の月より大きいものがなかったとき) */
    // 3列目の前に列を追加
    this.sheet.insertColumnBefore(3);

    // 追加した範囲を指定
    const monthRange = this.sheet.getRange(1, 3);

    // 範囲に値・フォントサイズ・配置・枠線を設定
    monthRange.setValue(month);
    monthRange.setFontSize(14);
    monthRange.setHorizontalAlignment("left");
    monthRange.setBorder(true, true, null, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    monthRange.setBorder(null, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);

    // 割合をいれる範囲を指定
    const dateRange = this.sheet.getRange(2, monthRange.getColumn());

    // 割合の範囲に値・フォントサイズ・配置・枠線を設定
    dateRange.setValue("割合");
    dateRange.setFontSize(10);
    dateRange.setHorizontalAlignment("center");
    dateRange.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);
    dateRange.setBorder(null, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    dateRange.setBackground(null);

    // その月のセルを返却
    return monthRange;
  }

  /**
   * 日のセルを挿入します(既に存在する場合は使用しないでください)
   * 月のセルがない場合は無効となります
   * @param {Date} ddate - Date型
   * @return {SpreadsheetApp.Range|null} 日のセル
   */
  insertDate(ddate) {
    const month = ddate.getMonth() + 1;
    const date = ddate.getDate();

    // 月のセルを取得
    const monthRange = this.getMonth(month);

    // 月のセルの列数を取得
    const monthNumColumns = monthRange.getNumColumns();

    // 月のセルの結合を解除
    monthRange.breakApart()

    /** 日が順番になるように追加する処理 */
    // 月のセルに既に日がある(=「範囲」以外のセルがある)場合
    if (monthRange.getNumColumns() - 1 > 0) {
      // 日にちの範囲を取得
      const dateRanges = this.sheet.getRange(2, monthRange.getColumn() + 1, 1, monthRange.getNumColumns() - 1);

      // 日にちの範囲の値を取得
      const dateValues = dateRanges.getValues()[0]

      // 日にちの値をループ
      for (let i in dateValues) {

        // 値が空だったらスキップ
        if (dateValues[i] == "") {
          continue;
        }

        // もし指定された日が値より小さかったら
        if (dateValues[i] < date) {

          // 後ろの日がある場合
          if (dateValues[Number(i) + 1]) {

            // 後ろの日の値が指定された日より小さかったらスキップ
            if (dateValues[Number(i) + 1] < date) {
              continue;
            }
          }

          /** それ以外の場合 */
          // 日のカラムの後ろに列を追加
          this.sheet.insertColumnAfter(Number(i) + monthRange.getColumn() + 1);

          // 追加した範囲を指定
          const dateRange = this.sheet.getRange(2, Number(i) + 5);

          // 範囲に値・フォントサイズ・配置を設定
          dateRange.setValue(date);
          dateRange.setFontSize(10);
          dateRange.setHorizontalAlignment("center");

          // 曜日を取得
          const day = ddate.getDay();

          // 祝日を取得
          const id = 'ja.japanese#holiday@group.v.calendar.google.com'
          const cal = CalendarApp.getCalendarById(id);
          const events = cal.getEventsForDay(ddate);

          if (day == 0) {
            dateRange.setBackground("#f4c7c3");
          } else if (day == 6) {
            dateRange.setBackground("#b7d2e1");
          } else if (events.length) {
            dateRange.setBackground("#c3e9f4");
          } else {
            dateRange.setBackground(null);
          }

          // 月のすべての日を取得
          const dates = this.sheet.getRange(2, monthRange.getColumn(), 1, monthNumColumns + 1);

          // その範囲に枠線を設定
          dates.setBorder(true, null, null, null, true, null, null, SpreadsheetApp.BorderStyle.SOLID);
          dates.setBorder(null, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

          // 月をセル結合
          this.sheet.getRange(1, monthRange.getColumn(), 1, monthNumColumns + 1).merge();

          // その日の出席情報が入るセルに枠線を設定
          const column = this.sheet.getRange(1, monthRange.getColumn(), this.sheet.getLastRow(), monthNumColumns + 1);
          column.setBorder(null, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
          column.setBorder(null, null, null, null, true, null, null, SpreadsheetApp.BorderStyle.SOLID);

          // その日のセルを返却
          return dateRange;
        }

      }
    }

    /** それ以外の場合(=指定された日が既存の日より大きいものがなかったとき) */
    // 割合の後ろに列を追加
    this.sheet.insertColumnAfter(monthRange.getColumn());

    // 追加した範囲を指定
    const dateRange = this.sheet.getRange(2, monthRange.getColumn() + 1);

    // 範囲に値・フォントサイズ・配置を設定
    dateRange.setValue(date);
    dateRange.setFontSize(10);
    dateRange.setHorizontalAlignment("center");

    // 曜日を取得
    const day = ddate.getDay();

    // 祝日を取得
    const id = 'ja.japanese#holiday@group.v.calendar.google.com'
    const cal = CalendarApp.getCalendarById(id);
    const events = cal.getEventsForDay(ddate);

    if (day == 0) {
      dateRange.setBackground("#f4c7c3");
    } else if (day == 6) {
      dateRange.setBackground("#b7d2e1");
    } else if (events.length) {
      dateRange.setBackground("#c3e9f4");
    } else {
      dateRange.setBackground(null);
    }

    // 月のすべての日を取得
    const dates = this.sheet.getRange(2, monthRange.getColumn(), 1, monthNumColumns + 1);

    // その範囲に枠線を設定
    dates.setBorder(true, null, null, null, true, null, null, SpreadsheetApp.BorderStyle.SOLID);
    dates.setBorder(null, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // 月をセル結合
    this.sheet.getRange(1, monthRange.getColumn(), 1, monthNumColumns + 1).merge();

    // その日の出席情報が入るセルに枠線を設定
    const column = this.sheet.getRange(2, monthRange.getColumn(), this.sheet.getLastRow() - 1, monthNumColumns + 1);
    column.setBorder(null, true, true, true, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    column.setBorder(null, null, null, null, true, null, null, SpreadsheetApp.BorderStyle.SOLID);

    // その日のセルを返却
    return dateRange;
  }

  /**
   * パートのセルを取得します
   * @param {String} part - パート
   * @return {SpreadsheetApp.Range|null} パートのセル(ない場合はnullを返却)
   */
  getPart(part) {
    // 1列目の範囲の値を取得
    const column = this.sheet.getRange(3, 1, this.sheet.getLastRow() - 2).getValues();

    // ループ
    for (let i in column) {

      // 値が指定されたパートだったら
      if (column[i][0] == part) {

        // その下の範囲の値を指定
        const belowValues = this.sheet.getRange(4 + Number(i), 1, this.sheet.getLastRow() - (3 + Number(i))).getValues();

        // 下の範囲をループ
        for (let j in belowValues) {

          // 値がいずれかのパートだったら
          for (let k in Part) {
            if (belowValues[j][0] == Part[k]) {

              // その1つ上のセルまでの範囲を返却
              return this.sheet.getRange(3 + Number(i), 1, Number(j) + 1);
            }
          }
        }

        // ない場合は一番下の1個上まで取得
        return this.sheet.getRange(`A${3 + Number(i)}:A${this.sheet.getLastRow() - 1}`);
      }
    }

    // 無い場合はnullを返却
    return null;
  }

  /**
   * すべてのパートのセルを取得します
   * @return {Object.<SpreadsheetApp.Range>} パートのセルのオブジェクト({"Fl": Range}のような形)
   */
  getAllPart() {
    // オブジェクトを初期化
    const partList = {}

    // パートでループ
    for (let i in Part) {

      // パートのセルを取得
      const partRange = this.getPart(Part[i]);

      // パートのセルがnullの場合はスキップ
      if (partRange == null) {
        continue;
      }

      // パートのセルをpartListに追加
      partList[i] = partRange;
    }

    // partListを返却
    return partList;
  }

  /**
   * 出席を記録します
   * @param {String} part - パート
   * @param {object} attendance - 出席情報
   * @param {Date} date - Date型
   */
  setAttendance(part, attendance, date) {

    /** いろいろ定義 */
    const dateRange = this.getDate(date)
    const monthRange = this.getMonth(date.getMonth() + 1);
    const monthColumn = monthRange.getColumn();
    const monthNumColumn = monthRange.getNumColumns();
    const partRange = this.getPart(part);
    const partNumRows = partRange.getNumRows();
    const partRow = partRange.getRow();
    const partValues = partRange.getValues();

    const lastColumn = this.sheet.getLastColumn();

    // 新しく追加した行を判定する用
    let newAdd = 0;

    // 行・列を参照する用
    let row = null;
    let column = null;

    /** 出席情報の処理 */
    // 出席情報をループ
    for (let i in attendance) {

      // 人のセルが存在したかを判定するための変数を指定
      let exists = false;

      // パートの値をループ
      for (let j in partValues) {

        // パートのセル内にその人の名前があったら
        if (partValues[j][0] == i) {

          // その人の名前の行と列を指定
          row = Number(j) + partRow;
          column = dateRange.getColumn();

          // 存在したのでtrueにする
          exists = true;
        }
      }

      // 人のセルがなかったら
      if (!exists) {

        // パートの最終列を取得
        const lastPartRow = partRange.getLastRow();

        // その最終列の後ろに行を追加
        this.sheet.insertRowAfter(lastPartRow + newAdd);

        // 新しく追加した行をカウント
        newAdd++;

        // その追加したセルを指定
        const nameRange = this.sheet.getRange(lastPartRow + newAdd, 1);

        // そのセルの高さ・フォントサイズ・値を設定
        this.sheet.setRowHeight(lastPartRow + newAdd, 10);
        nameRange.setFontSize(10);
        nameRange.setValue(i);

        // 行・列を設定
        row = lastPartRow + newAdd;
        column = dateRange.getColumn();
      }

      // 出席情報を入れるセルを指定
      const attendanceRange = this.sheet.getRange(row, column);

      // そのセルに出席情報の値を設定
      attendanceRange.setValue(attendance[i]);

      // その人の行を指定
      const nameRow = this.sheet.getRange(row, 1, 1, lastColumn);

      // 行に枠線を設定
      nameRow.setBorder(true, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID);

      // もしその人がパートの最終行だったら下を太い枠線にする
      if (row == Object.keys(attendance).length + partRow) {
        nameRow.setBorder(null, null, true, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      }

      // もしその人がパートの最初の行だったら上を二重線にする
      if (row == partRow + 1) {
        nameRow.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.DOUBLE);
      }

      // その人の月の平均を設定
      this.sheet.getRange(row, monthColumn)
        .setValue(`=AttendanceAverage(${this.sheet.getRange(row, monthColumn + 1, 1, monthNumColumn - 1).getA1Notation()})`)
        .setHorizontalAlignment("right");

      // その人の全体の平均を設定
      this.sheet.getRange(row, 2)
        .setValue(`=AttendanceAverage(${this.sheet.getRange(row, 3, 1, lastColumn - 2).getA1Notation()})`)
        .setHorizontalAlignment("right");
    }

    /** 最終行取得 */
    const lastRow = this.sheet.getLastRow();

    // パート名の上に枠線を設定
    this.sheet.getRange(partRange.getRow(), 1, 1, lastColumn).setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    /** 平均などの処理 */
    // 日の平均を設定
    const dRange = this.sheet.getRange(lastRow, column);
    dRange.setValue(`=AttendanceAverage(${this.sheet.getRange(3, column, lastRow - 3).getA1Notation()})`);
    dRange.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    dRange.setHorizontalAlignment("right");

    // 月の平均を設定
    const mRange = this.sheet.getRange(lastRow, monthColumn);
    mRange.setValue(`=AttendanceAverage(${this.sheet.getRange(3, monthColumn + 1, lastRow - 3, monthNumColumn - 1).getA1Notation()})`);
    mRange.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    mRange.setHorizontalAlignment("right");

    // 全体の平均を設定
    const wRange = this.sheet.getRange(lastRow, 2);
    wRange.setValue(`=AttendanceAverage(${this.sheet.getRange(3, 3, lastRow - 3, lastColumn - 2).getA1Notation()})`);
    wRange.setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    wRange.setHorizontalAlignment("right");

    // パートの日の平均を設定
    const pdRange = this.sheet.getRange(partRow, column);
    pdRange.setValue(`=AttendanceAverage(${this.sheet.getRange(partRow + 1, column, partNumRows + newAdd - 1).getA1Notation()})`);
    pdRange.setHorizontalAlignment("right");

    // パートの月の平均を設定
    const pmRange = this.sheet.getRange(partRow, monthColumn);
    pmRange.setValue(`=AttendanceAverage(${this.sheet.getRange(partRow + 1, monthColumn + 1, partNumRows + newAdd - 1, monthNumColumn - 1).getA1Notation()})`);
    pmRange.setHorizontalAlignment("right");

    // パートの全体の平均を設定
    const pRange = this.sheet.getRange(partRow, 2);
    pRange.setValue(`=AttendanceAverage(${this.sheet.getRange(partRow + 1, 3, partNumRows + newAdd - 1, lastColumn - 2).getA1Notation()})`);
    pRange.setHorizontalAlignment("right");
  }
}
