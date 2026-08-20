/**
 * ============================================================
 * 技術室 工作機械予約システム
 * Google Apps Script版
 *
 * 保存先：
 *   Googleスプレッドシート
 *
 * 主な機能：
 *   ・機械種類・台数管理
 *   ・授業時間管理
 *   ・最大使用時間管理
 *   ・生徒予約
 *   ・予約重複防止
 *   ・進捗管理
 *   ・教師設定
 *   ・CSV出力用データ取得
 *   ・大型モニター表示
 *
 * 外部ライブラリ：
 *   使用していません。
 * ============================================================
 */


/* ============================================================
 * 1. 基本設定
 * ============================================================
 */

/**
 * このGASをスプレッドシートから
 * 「拡張機能 → Apps Script」で作成した場合、
 * 空欄のままでOKです。
 *
 * スタンドアロン型GASの場合は、
 * 対象スプレッドシートのIDを入力してください。
 */
const SPREADSHEET_ID = "";


/**
 * シート名
 */
const SHEET_RESERVATIONS = "予約";
const SHEET_MACHINES = "機械";
const SHEET_SETTINGS = "設定";


/**
 * 初期教師PIN
 *
 * 初回設定後、教師設定画面から変更してください。
 *
 * 実際のPINはScript Propertiesに保存します。
 */
const DEFAULT_TEACHER_PIN = "1234";


/**
 * 予約時間の刻み
 *
 * 5分刻み
 */
const TIME_STEP_MINUTES = 5;


/* ============================================================
 * 2. Webアプリ起動
 * ============================================================
 */

/**
 * Webアプリの入口
 *
 * URL例：
 *
 * 生徒用
 * ?mode=student
 *
 * 大型モニター用
 * ?mode=display
 *
 * 教師設定
 * ?mode=teacher
 */
function doGet(e) {

  initializeSystem();

  const template =
    HtmlService.createTemplateFromFile("Index");

  const mode =
    e && e.parameter && e.parameter.mode
      ? e.parameter.mode
      : "student";

  template.mode = mode;

  return template
    .evaluate()
    .setTitle("技術室 工作機械予約システム")
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}


/* ============================================================
 * 3. スプレッドシート取得
 * ============================================================
 */

function getSpreadsheet() {

  if (SPREADSHEET_ID) {

    return SpreadsheetApp.openById(
      SPREADSHEET_ID
    );

  }


  const active =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!active) {

    throw new Error(
      "対象スプレッドシートを取得できません。"
    );

  }

  return active;
}


/* ============================================================
 * 4. 初期化
 * ============================================================
 */

/**
 * 必要なシートを自動作成する
 */
function initializeSystem() {

  const ss =
    getSpreadsheet();


  /*
   * 予約シート
   */
  let reservationSheet =
    ss.getSheetByName(
      SHEET_RESERVATIONS
    );


  if (!reservationSheet) {

    reservationSheet =
      ss.insertSheet(
        SHEET_RESERVATIONS
      );

    reservationSheet.appendRow([

      "ID",
      "実施日",
      "名簿番号",
      "機械ID",
      "機械名",
      "台",
      "開始時刻",
      "終了時刻",
      "進捗状況",
      "登録日時"

    ]);

    reservationSheet
      .getRange(1, 1, 1, 10)
      .setFontWeight("bold");

  }


  /*
   * 機械シート
   */
  let machineSheet =
    ss.getSheetByName(
      SHEET_MACHINES
    );


  if (!machineSheet) {

    machineSheet =
      ss.insertSheet(
        SHEET_MACHINES
      );

    machineSheet.appendRow([

      "機械ID",
      "機械名",
      "台数"

    ]);

    machineSheet.appendRow([
      "drill",
      "ボール盤",
      2
    ]);

    machineSheet.appendRow([
      "scroll",
      "糸のこ盤",
      2
    ]);

    machineSheet.appendRow([
      "sander",
      "ベルトサンダー",
      1
    ]);

    machineSheet
      .getRange(1, 1, 1, 3)
      .setFontWeight("bold");

  }


  /*
   * 設定シート
   */
  let settingsSheet =
    ss.getSheetByName(
      SHEET_SETTINGS
    );


  if (!settingsSheet) {

    settingsSheet =
      ss.insertSheet(
        SHEET_SETTINGS
      );

    settingsSheet.appendRow([
      "項目",
      "値"
    ]);

    settingsSheet.appendRow([
      "授業開始",
      "13:00"
    ]);

    settingsSheet.appendRow([
      "授業終了",
      "14:30"
    ]);

    settingsSheet.appendRow([
      "最大使用時間",
      "20"
    ]);

    settingsSheet
      .getRange(1, 1, 1, 2)
      .setFontWeight("bold");

  }


  /*
   * 教師PIN
   */
  const properties =
    PropertiesService.getScriptProperties();

  if (
    !properties.getProperty(
      "TEACHER_PIN"
    )
  ) {

    properties.setProperty(
      "TEACHER_PIN",
      DEFAULT_TEACHER_PIN
    );

  }

}


/* ============================================================
 * 5. 時刻処理
 * ============================================================
 */

/**
 * HH:mm → 分
 */
function timeToMinutes(time) {

  if (!time) {
    return null;
  }

  const parts =
    String(time).split(":");

  return (
    Number(parts[0]) * 60 +
    Number(parts[1])
  );
}


/**
 * 分 → HH:mm
 */
function minutesToTime(minutes) {

  const hour =
    Math.floor(minutes / 60);

  const minute =
    minutes % 60;

  return (
    String(hour).padStart(2, "0") +
    ":" +
    String(minute).padStart(2, "0")
  );
}


/* ============================================================
 * 6. 今日の日付
 * ============================================================
 */

function getTodayString() {

  const timezone =
    Session.getScriptTimeZone() ||
    "Asia/Tokyo";

  return Utilities.formatDate(
    new Date(),
    timezone,
    "yyyy-MM-dd"
  );
}


/**
 * 日付＋時刻
 */
function getNowString() {

  const timezone =
    Session.getScriptTimeZone() ||
    "Asia/Tokyo";

  return Utilities.formatDate(
    new Date(),
    timezone,
    "yyyy-MM-dd HH:mm:ss"
  );
}


/* ============================================================
 * 7. 設定取得
 * ============================================================
 */

function getSystemData() {

  initializeSystem();

  return {

    settings:
      getSettings(),

    machines:
      getMachines(),

    reservations:
      getReservationsForDate(
        getTodayString()
      )

  };
}


/**
 * 設定取得
 */
function getSettings() {

  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_SETTINGS
      );


  const values =
    sheet.getDataRange()
      .getValues();


  const settings = {

    lessonStart: "13:00",
    lessonEnd: "14:30",
    maxMinutes: 20

  };


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const key =
      String(values[i][0]);

    const value =
      String(values[i][1]);


    if (
      key === "授業開始"
    ) {

      settings.lessonStart =
        value;

    }


    if (
      key === "授業終了"
    ) {

      settings.lessonEnd =
        value;

    }


    if (
      key === "最大使用時間"
    ) {

      settings.maxMinutes =
        Number(value);

    }

  }


  return settings;
}


/**
 * 機械一覧取得
 */
function getMachines() {

  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_MACHINES
      );


  const values =
    sheet.getDataRange()
      .getValues();


  const machines = [];


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      !values[i][0] ||
      !values[i][1]
    ) {
      continue;
    }


    machines.push({

      id:
        String(values[i][0]),

      name:
        String(values[i][1]),

      count:
        Number(values[i][2]) || 1

    });

  }


  return machines;
}


/* ============================================================
 * 8. 予約取得
 * ============================================================
 */

/**
 * 指定日の予約を取得
 */
function getReservationsForDate(
  date
) {

  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_RESERVATIONS
      );


  const values =
    sheet.getDataRange()
      .getValues();


  const reservations = [];


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    const row =
      values[i];


    if (
      !row[0]
    ) {
      continue;
    }


    const rowDate =
      normalizeDateValue(
        row[1]
      );


    if (
      rowDate !== date
    ) {
      continue;
    }


    const reservation = {

      id:
        String(row[0]),

      date:
        rowDate,

      studentNumber:
        String(row[2]),

      machineId:
        String(row[3]),

      machineName:
        String(row[4]),

      unit:
        String(row[5]),

      startTime:
        normalizeTimeValue(
          row[6]
        ),

      endTime:
        normalizeTimeValue(
          row[7]
        ),

      status:
        String(row[8]),

      createdAt:
        String(row[9])

    };


    /*
     * 現在時刻から進捗を再計算
     */
    reservation.status =
      calculateStatus(
        reservation
      );


    reservations.push(
      reservation
    );

  }


  return reservations;
}


/**
 * 日付値を文字列に変換
 */
function normalizeDateValue(
  value
) {

  if (
    value instanceof Date
  ) {

    const timezone =
      Session.getScriptTimeZone() ||
      "Asia/Tokyo";

    return Utilities.formatDate(
      value,
      timezone,
      "yyyy-MM-dd"
    );

  }


  const text =
    String(value);


  if (
    text.includes("T")
  ) {

    return text.substring(
      0,
      10
    );

  }


  return text;
}


/**
 * 時刻値を文字列に変換
 */
function normalizeTimeValue(
  value
) {

  if (
    value instanceof Date
  ) {

    const timezone =
      Session.getScriptTimeZone() ||
      "Asia/Tokyo";

    return Utilities.formatDate(
      value,
      timezone,
      "HH:mm"
    );

  }


  return String(value);
}


/* ============================================================
 * 9. 進捗状況
 * ============================================================
 */

function calculateStatus(
  reservation
) {

  const now =
    getCurrentMinutes();


  const start =
    timeToMinutes(
      reservation.startTime
    );


  const end =
    timeToMinutes(
      reservation.endTime
    );


  if (
    reservation.status ===
    "キャンセル"
  ) {

    return "キャンセル";

  }


  if (
    now < start
  ) {

    return "予約前";

  }


  if (
    now >= start &&
    now < end
  ) {

    return "使用中";

  }


  return "完了";
}


/**
 * 現在時刻を分で取得
 */
function getCurrentMinutes() {

  const timezone =
    Session.getScriptTimeZone() ||
    "Asia/Tokyo";

  const time =
    Utilities.formatDate(
      new Date(),
      timezone,
      "HH:mm"
    );

  return timeToMinutes(
    time
  );
}


/* ============================================================
 * 10. 予約登録
 * ============================================================
 */

/**
 * 生徒からの予約登録
 */
function createReservation(
  data
) {

  initializeSystem();


  /*
   * 同時アクセス対策
   *
   * 同じ機械を複数人が同時に予約しても、
   * 同時処理にならないようにする。
   */
  const lock =
    LockService.getScriptLock();


  lock.waitLock(
    10000
  );


  try {

    const result =
      validateReservation(
        data
      );


    if (!result.ok) {

      return {

        success: false,

        message:
          result.message

      };

    }


    const sheet =
      getSpreadsheet()
        .getSheetByName(
          SHEET_RESERVATIONS
        );


    const reservationId =
      "R-" +
      Utilities.getUuid();


    const now =
      getNowString();


    sheet.appendRow([

      reservationId,

      data.date,

      String(data.studentNumber),

      data.machineId,

      data.machineName,

      String(data.unit),

      data.startTime,

      data.endTime,

      "予約前",

      now

    ]);


    return {

      success: true,

      message:
        "予約しました。",

      reservation: {

        id:
          reservationId,

        date:
          data.date,

        studentNumber:
          String(
            data.studentNumber
          ),

        machineId:
          data.machineId,

        machineName:
          data.machineName,

        unit:
          String(data.unit),

        startTime:
          data.startTime,

        endTime:
          data.endTime,

        status:
          "予約前",

        createdAt:
          now

      }

    };


  } finally {

    lock.releaseLock();

  }

}


/* ============================================================
 * 11. 予約チェック
 * ============================================================
 */

function validateReservation(
  data
) {

  if (!data) {

    return {
      ok: false,
      message:
        "予約データがありません。"
    };

  }


  /*
   * 名簿番号
   */
  if (
    !String(
      data.studentNumber || ""
    ).trim()
  ) {

    return {
      ok: false,
      message:
        "名簿番号を入力してください。"
    };

  }


  /*
   * 機械
   */
  const machines =
    getMachines();


  const machine =
    machines.find(
      m =>
        m.id ===
        data.machineId
    );


  if (!machine) {

    return {
      ok: false,
      message:
        "選択された機械が見つかりません。"
    };

  }


  const unit =
    Number(data.unit);


  if (
    !unit ||
    unit < 1 ||
    unit > machine.count
  ) {

    return {
      ok: false,
      message:
        "機械の台番号が正しくありません。"
    };

  }


  /*
   * 時間
   */
  const start =
    timeToMinutes(
      data.startTime
    );


  const end =
    timeToMinutes(
      data.endTime
    );


  if (
    start === null ||
    end === null
  ) {

    return {
      ok: false,
      message:
        "使用時間を入力してください。"
    };

  }


  if (
    end <= start
  ) {

    return {
      ok: false,
      message:
        "終了時刻は開始時刻より後にしてください。"
    };

  }


  const settings =
    getSettings();


  const lessonStart =
    timeToMinutes(
      settings.lessonStart
    );


  const lessonEnd =
    timeToMinutes(
      settings.lessonEnd
    );


  /*
   * 授業時間内
   */
  if (
    start < lessonStart ||
    end > lessonEnd
  ) {

    return {
      ok: false,
      message:
        `授業時間 ${settings.lessonStart}～${settings.lessonEnd} の範囲内で予約してください。`
    };

  }


  /*
   * 最大時間
   */
  const duration =
    end - start;


  if (
    duration >
    Number(settings.maxMinutes)
  ) {

    return {
      ok: false,
      message:
        `1回の最大使用時間は${settings.maxMinutes}分です。`
    };

  }


  /*
   * 5分刻み
   */
  if (
    start % TIME_STEP_MINUTES !== 0 ||
    end % TIME_STEP_MINUTES !== 0
  ) {

    return {
      ok: false,
      message:
        "予約時刻は5分刻みで指定してください。"
    };

  }


  /*
   * 今日の日付か確認
   */
  const today =
    getTodayString();


  if (
    data.date !== today
  ) {

    return {
      ok: false,
      message:
        "本日の授業について予約してください。"
    };

  }


  /*
   * 過去時間を禁止
   */
  const now =
    getCurrentMinutes();


  if (
    start < now
  ) {

    return {
      ok: false,
      message:
        "現在時刻より前の時間は予約できません。"
    };

  }


  /*
   * 同一機械・同一台の重複チェック
   */
  const reservations =
    getReservationsForDate(
      today
    );


  const duplicate =
    reservations.some(
      reservation => {

        if (
          reservation.status ===
          "キャンセル"
        ) {
          return false;
        }


        if (
          reservation.machineId !==
          data.machineId
        ) {
          return false;
        }


        if (
          String(
            reservation.unit
          ) !==
          String(unit)
        ) {
          return false;
        }


        const rStart =
          timeToMinutes(
            reservation.startTime
          );


        const rEnd =
          timeToMinutes(
            reservation.endTime
          );


        /*
         * 時間重複判定
         */
        return (
          start < rEnd &&
          end > rStart
        );

      }
    );


  if (duplicate) {

    return {
      ok: false,
      message:
        "この機械・この台は、その時間にすでに予約があります。"
    };

  }


  /*
   * 同じ名簿番号の同時間帯予約も禁止
   *
   * 1人が同時に2台を予約することを防ぐ。
   */
  const studentDuplicate =
    reservations.some(
      reservation => {

        if (
          reservation.status ===
          "キャンセル"
        ) {
          return false;
        }


        if (
          String(
            reservation.studentNumber
          ) !==
          String(
            data.studentNumber
          )
        ) {
          return false;
        }


        const rStart =
          timeToMinutes(
            reservation.startTime
          );


        const rEnd =
          timeToMinutes(
            reservation.endTime
          );


        return (
          start < rEnd &&
          end > rStart
        );

      }
    );


  if (studentDuplicate) {

    return {
      ok: false,
      message:
        "同じ時間帯に別の機械を予約しています。"
    };

  }


  return {

    ok: true

  };

}


/* ============================================================
 * 12. 予約キャンセル
 * ============================================================
 */

function cancelReservation(
  reservationId
) {

  const lock =
    LockService.getScriptLock();


  lock.waitLock(
    10000
  );


  try {

    const sheet =
      getSpreadsheet()
        .getSheetByName(
          SHEET_RESERVATIONS
        );


    const values =
      sheet.getDataRange()
        .getValues();


    for (
      let i = 1;
      i < values.length;
      i++
    ) {

      if (
        String(values[i][0]) ===
        String(reservationId)
      ) {

        /*
         * 9列目 = 進捗状況
         */
        sheet
          .getRange(
            i + 1,
            9
          )
          .setValue(
            "キャンセル"
          );


        return {

          success: true,

          message:
            "予約をキャンセルしました。"

        };

      }

    }


    return {

      success: false,

      message:
        "予約が見つかりません。"

    };


  } finally {

    lock.releaseLock();

  }

}


/* ============================================================
 * 13. 教師認証
 * ============================================================
 */

/**
 * 教師PINを確認
 */
function verifyTeacherPin(
  pin
) {

  const properties =
    PropertiesService
      .getScriptProperties();


  const savedPin =
    properties.getProperty(
      "TEACHER_PIN"
    );


  return {

    success:
      String(pin) ===
      String(savedPin)

  };

}


/**
 * 教師PIN変更
 */
function changeTeacherPin(
  oldPin,
  newPin
) {

  const check =
    verifyTeacherPin(
      oldPin
    );


  if (!check.success) {

    return {

      success: false,

      message:
        "現在の教師PINが違います。"

    };

  }


  if (
    !newPin ||
    String(newPin).length < 4
  ) {

    return {

      success: false,

      message:
        "PINは4桁以上にしてください。"

    };

  }


  PropertiesService
    .getScriptProperties()
    .setProperty(
      "TEACHER_PIN",
      String(newPin)
    );


  return {

    success: true,

    message:
      "教師PINを変更しました。"

  };

}


/* ============================================================
 * 14. 教師設定保存
 * ============================================================
 */

function saveTeacherSettings(
  pin,
  data
) {

  const auth =
    verifyTeacherPin(
      pin
    );


  if (!auth.success) {

    return {

      success: false,

      message:
        "教師PINが違います。"

    };

  }


  /*
   * 入力チェック
   */
  const lessonStart =
    timeToMinutes(
      data.lessonStart
    );


  const lessonEnd =
    timeToMinutes(
      data.lessonEnd
    );


  if (
    lessonStart === null ||
    lessonEnd === null ||
    lessonEnd <= lessonStart
  ) {

    return {

      success: false,

      message:
        "授業時間が正しくありません。"

    };

  }


  const maxMinutes =
    Number(
      data.maxMinutes
    );


  if (
    !maxMinutes ||
    maxMinutes < 1
  ) {

    return {

      success: false,

      message:
        "最大使用時間が正しくありません。"

    };

  }


  /*
   * 設定シート
   */
  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_SETTINGS
      );


  const values =
    sheet.getDataRange()
      .getValues();


  const settingMap = {

    "授業開始":
      data.lessonStart,

    "授業終了":
      data.lessonEnd,

    "最大使用時間":
      maxMinutes

  };


  Object.keys(
    settingMap
  ).forEach(
    key => {

      let found =
        false;


      for (
        let i = 1;
        i < values.length;
        i++
      ) {

        if (
          String(values[i][0]) ===
          key
        ) {

          sheet
            .getRange(
              i + 1,
              2
            )
            .setValue(
              settingMap[key]
            );

          found = true;

          break;

        }

      }


      if (!found) {

        sheet.appendRow([
          key,
          settingMap[key]
        ]);

      }

    }
  );


  return {

    success: true,

    message:
      "授業設定を保存しました。"

  };

}


/* ============================================================
 * 15. 機械設定保存
 * ============================================================
 */

function saveMachineSettings(
  pin,
  machines
) {

  const auth =
    verifyTeacherPin(
      pin
    );


  if (!auth.success) {

    return {

      success: false,

      message:
        "教師PINが違います。"

    };

  }


  if (
    !Array.isArray(machines)
  ) {

    return {

      success: false,

      message:
        "機械データが正しくありません。"

    };

  }


  /*
   * 入力チェック
   */
  const cleaned = [];


  machines.forEach(
    machine => {

      const name =
        String(
          machine.name || ""
        ).trim();

      const count =
        Number(
          machine.count
        );


      if (!name) {
        return;
      }


      if (
        !count ||
        count < 1
      ) {
        return;
      }


      cleaned.push({

        id:
          String(
            machine.id ||
            "machine_" +
            Utilities.getUuid()
          ),

        name:
          name,

        count:
          count

      });

    }
  );


  /*
   * シートを更新
   */
  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_MACHINES
      );


  sheet.clearContents();


  sheet.appendRow([
    "機械ID",
    "機械名",
    "台数"
  ]);


  cleaned.forEach(
    machine => {

      sheet.appendRow([

        machine.id,

        machine.name,

        machine.count

      ]);

    }
  );


  return {

    success: true,

    message:
      "機械設定を保存しました。",

    machines:
      cleaned

  };

}


/* ============================================================
 * 16. 全予約削除
 * ============================================================
 */

function deleteTodayReservations(
  pin
) {

  const auth =
    verifyTeacherPin(
      pin
    );


  if (!auth.success) {

    return {

      success: false,

      message:
        "教師PINが違います。"

    };

  }


  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_RESERVATIONS
      );


  const values =
    sheet.getDataRange()
      .getValues();


  const today =
    getTodayString();


  /*
   * 下から削除
   */
  for (
    let i = values.length - 1;
    i >= 1;
    i--
  ) {

    const date =
      normalizeDateValue(
        values[i][1]
      );


    if (
      date === today
    ) {

      sheet.deleteRow(
        i + 1
      );

    }

  }


  return {

    success: true,

    message:
      "本日の予約をすべて削除しました。"

  };

}


/* ============================================================
 * 17. CSV用データ
 * ============================================================
 */

/**
 * 保存データをすべて取得
 */
function getAllReservations() {

  const sheet =
    getSpreadsheet()
      .getSheetByName(
        SHEET_RESERVATIONS
      );


  const values =
    sheet.getDataRange()
      .getValues();


  const data = [];


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (!values[i][0]) {
      continue;
    }


    data.push({

      id:
        String(values[i][0]),

      date:
        normalizeDateValue(
          values[i][1]
        ),

      studentNumber:
        String(values[i][2]),

      machineId:
        String(values[i][3]),

      machineName:
        String(values[i][4]),

      unit:
        String(values[i][5]),

      startTime:
        normalizeTimeValue(
          values[i][6]
        ),

      endTime:
        normalizeTimeValue(
          values[i][7]
        ),

      status:
        String(values[i][8]),

      createdAt:
        String(values[i][9])

    });

  }


  return data;
}