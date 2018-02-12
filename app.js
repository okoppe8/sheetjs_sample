var X = XLSX;

// メイン処理
function handleFile(e) {

    var files = e.target.files;
    var f = files[0]; {
        var reader = new FileReader();
        reader.onload = function (e) {
            var data = e.target.result;
            var wb;
            var arr = fixdata(data);
            wb = X.read(btoa(arr), {
                type: 'base64',
                cellDates: true,
                dateNF: 'yyyy/mm/dd;@'
            });

            var output = "";
            output = to_json(wb);

            console.log(output);
            $("pre#result").html(JSON.stringify(output, null, 2));

        };
        reader.readAsArrayBuffer(f);
    }
}
// エクセルファイルを読み込み
function fixdata(data) {
    var o = "",
        l = 0,
        w = 10240;
    for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w,
        l * w + w)));
    o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
    return o;
}
// ワークブックのデータをjsonに変換
function to_json(workbook) {
    var result = {};
    workbook.SheetNames.forEach(function (sheetName) {
        var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName],{raw:true, header:0});
        if (roa.length > 0) {
            result[sheetName] = roa;
        }
    });
    return result;
}

// 画面初期化
$(document).ready(function () {

    // // ファイル選択欄 選択イベント
    $('.custom-file-input').on('change', function (e) {
        handleFile(e);
        $(this).next('.custom-file-label').html($(this)[0].files[0].name);
    })

});

