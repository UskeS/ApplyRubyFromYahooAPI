//-- プロトタイプ拡張 --//
if (!String.prototype.surroundQuotes) {
    String.prototype.surroundQuotes = function(q) {
        return q + this + q;
    }
}

//-- 関数定義 --//
/**
 * オブジェクトからパラメータ用の文字列を生成する関数
 * 型によって引用符で囲むかどうか変える
 * Object型が渡されたら再帰的に解決し {key: value} の形にする
 */
function getParamString(obj) {
    var res = [];
    for (var key in obj) {
        if (typeof(obj[key]) === "string") {
            res.push(key.surroundQuotes('\\"') + ": " + obj[key].surroundQuotes('\\"'));
        } else if (typeof(obj[key]) === "number") {
            res.push(key.surroundQuotes('\\"') + ": " + obj[key]);
        } else {
            res.push(key.surroundQuotes('\\"') + ": " + getParamString(obj[key]));
        }
    }
    return "{" + res.join(",") + "}";
}

/**
 * Yahoo! テキスト解析APIを利用するための設定を読み込む関数。
 * config.jsonファイルがスクリプトと同階層になければ生成する。
 * config.jsonには`clientID`と`grade`を書き込む。
 * 可変にしたい設定値があればこちらに移動させる。
 * 初回実行時はconfig.jsonを生成してスクリプトは一旦終了する（clientIDを書き換えてもらう必要がある）。
 * 何か問題があった場合は`false`を返す。
 * 問題なく設定が読み込めたらJSONを返す。
 *
 * @return {Object} clientIDとgradeをもったObject型オブジェクト
 */
function getConfig() {
    var myConfig = {};
    var filepath = decodeURIComponent(File($.fileName).parent) + "/config.json";
    var configFile = File(filepath);
    if (!configFile.exists) {
        var defaultConfigString = '{ "clientID": "xxxxxxxxxx", "grade": 1 }';
        var openFlag = configFile.open("w");
        configFile.encoding = "UTF-8";
        var er;
        if (!openFlag) {
            alert("config.jsonがスクリプトと同じ階層に生成できません。アクセス権限等を確認してください。");
            exit();
        }
        try {
            configFile.write(defaultConfigString);
        } catch (e) {
            er = e;
        } finally {
            configFile.close();
        }
        if (er) {
            alert(er);
            exit();
        }
        alert("スクリプトと同じ階層にconfig.jsonという設定ファイルを生成しました。\rテキストエディタで開き、xxxxxxxxxx のところに、取得したYahoo! APIのアプリケーションIDを記述してください");
        return false;
    } else {
        myConfig = function() {
            var temp, er, res;
            var openFlag = configFile.open("r");
            if (!openFlag) { return false; }
            openFlag.encoding = "UTF-8";
            try {
                temp = configFile.read(99999);
            } catch (e) {
                er = e;
            } finally {
                configFile.close();
            }
            if (er) {
                alert(er);
                return false;
            }
            try {
                res = eval("(" + temp + ")");
            } catch (e) {
                alert(e);
                return false;
            }
            return res;
        }();
    }
    if (!myConfig) {
        alert("config.jsonが読み込めませんでした。config.jsonファイルを削除してから再度実行してください。");
        exit();
    }
    return myConfig;
}

/**
 * 渡されたJSONの値の型を確認する関数
 *
 * @param {Object} obj チェックしたいObject型オブジェクト
 * @return {Boolean} 
 */
function validateJSON(obj) {
    var erString = ["config.jsonの内容に不具合があります。"];
    var templateString = "★★★プロパティは※※※型である必要があります。";
    var validateList = {
        clientID: "String",
        grade: "Number",
    };
    for (var k in obj) {
        if (obj[k].constructor.name !== validateList[k]) {
            var actString = templateString.replace("★★★", k).replace("※※※", validateList[k]);
            erString.push(actString);
        }
    }
    if (erString.length > 1) {
        alert(erString.join("\r"));
        return false;
    }
    return true;
}

//-- 実行処理 ここから --//
var myConf = getConfig();
if (!myConf || !validateJSON(myConf)) {
    alert("スクリプトを終了します");
    exit();
}

/** 
 * スクリプト実行前の状況確認を行う即時関数
 */
! function() {
    var mes = "";
    if (app.documents.length === 0) {
        mes = "ドキュメントを開いてください";
    } else if (app.activeDocument.selection.length !== 1 ||
        !app.activeDocument.selection[0].hasOwnProperty("contents")) {
        mes = "テキストを選択してください";
    }
    if (mes) {
        alert(mes);
        exit();
    }
}();

var doc = app.activeDocument;
var sel = doc.selection[0];

/** 
 * リクエストの結果を格納する変数。
 * 例外が起きると error プロパティが生成される。
 * 成功すると result プロパティが生成され、親文字とルビの語句がまとめて取得できるようになる。
 * レスポンスフィールドは公式ドキュメント参照：
 * https://developer.yahoo.co.jp/webapi/jlp/furigana/v2/furigana.html
 */
var response = {};

/**
 * Yahoo! のテキスト解析APIにpostリクエストする即時関数
 */
! function() {
    var p = {
        url: "https://jlp.yahooapis.jp/FuriganaService/V2/furigana",
        header: {
            contentType: "application/json",
            userAgent: "Yahoo AppID: " + myConf.clientID,
        },
        body: {
            id: "1234",
            jsonrpc: "2.0",
            method: "jlp.furiganaservice.furigana",
            params: {
                q: sel.contents,
                grade: myConf.grade,
            }
        },
    };

    var scpt = "do shell script " + [
        "curl",
        "-f",
        "-H",
        ("content-type: " + p.header.contentType).surroundQuotes("'"),
        "-A",
        p.header.userAgent.surroundQuotes("'"),
        "-X",
        "POST",
        "-d",
        getParamString(p.body).surroundQuotes("'"),
        p.url
    ].join(" ").surroundQuotes('"');

    var isError = false;

    try {
        response = eval("(" + app.doScript(scpt, ScriptLanguage.applescriptLanguage) + ")");
    } catch (e) {
        alert("#" + $.line + "; 下記のエラーが発生しました。Yahoo! テキスト解析APIのドキュメントを参照してください\r" + e.name + ": \r" + e.message);
        // エラーコードは右記参照 https://developer.yahoo.co.jp/appendix/errors.html
        isError = true;
    } finally {
        p,
        scpt,
        isError = null;
    }
    if (isError) { exit(); }
}();

/**
 * response 変数の中身をチェック
 * ここがクリアされれば問題なくルビが格納されている
 */
if (response.error) {
    var er = response.error;
    alert("#" + $.line + "; 下記のエラーが発生しました。エラーコードはJSON-RPC 2.0 #Response objectを参照してください\rcode: " + er.code + ", \rmessage: " + er.message);
    // JSON-RPC 2.0 #Response object https://www.jsonrpc.org/specification#response_object
    exit();
} else if (!response.result) {
    alert("#" + $.line + "; 結果が取得できませんでした");
    exit();
}

