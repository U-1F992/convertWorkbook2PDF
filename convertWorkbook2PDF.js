function main() {

    forceCScript(WScript.Arguments);

    WScript.StdOut.WriteLine("Excel Workbook => PDF converter\n")

    checkArgs(WScript.Arguments);

    feedArgs2Converter(WScript.Arguments);

    WScript.Quit(0);
}

function forceCScript(args) {
    // CScriptでの起動を強制する
    // 引数も引き継いで渡す
    if (WScript.FullName.slice(-"*Script.exe".length).toLowerCase() == "wscript.exe") {

        var str = "";
        for (var i = 0; i < args.Count(); i++) {
            str += " \"" + args.Item(i) + "\""
        }
        
        new ActiveXObject("WScript.Shell").Run("cscript \"" + WScript.ScriptFullName + "\"" + str);
        WScript.Quit(0);
    }
}

function checkArgs(args) {
    // 引数チェック
    if (args.length == 0) {
        // 引数がない場合
        WScript.StdErr.WriteLine("Please specify target with arguments.");
        WScript.Quit(1);

    } else {
        var fso = new ActiveXObject("Scripting.FileSystemObject");
        
        for(var i = 0; i < args.Count(); i++) {
            if (fso.folderExists(args.Item(i)) == false 
            && !boolSupportedFormat(args.Item(i))) {
                
                // 引数がフォルダではないかつワークブックではないファイルを指している場合
                WScript.StdErr.WriteLine(args.Item(i) + " is not supported format.");
                WScript.Quit(1);

            } else if (fso.folderExists(args.Item(i)) == true) {

                // 引数がフォルダではあるが中にワークブックがない場合
                var flag = false;
                var objFolder = fso.getFolder(WScript.Arguments.Item(i) + "\\");
                var e = new Enumerator(objFolder.Files);
                for (e.moveFirst(); !e.atEnd(); e.moveNext()) {
                    if (boolSupportedFormat(e.item().name)) {
                        flag = true;
                        break;                        
                    }
                }
                if (flag == false) {
                    WScript.StdErr.WriteLine("There is not any workbook in " + args.Item(i) + ".");
                    WScript.Quit(1);
                }
            }
        }
    }
}

function boolSupportedFormat (strFileName) {

    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var strExtension = fso.getExtensionName(strFileName).toLowerCase();

    if ( strExtension == "xlsx"
      || strExtension == "xlsm"
      || strExtension == "xls"
      || strExtension == "csv") {
        return true;
    } else {
        return false;
    }
}

function feedArgs2Converter (args) {

    var fso = new ActiveXObject("Scripting.FileSystemObject");
    for (var i = 0; i < args.Count(); i++) {

        if (fso.folderExists(args.Item(i)) == false) {
            // ファイルの場合
            WScript.StdOut.WriteLine("Processing : " + args.Item(i));
            convertWorkbook2PDF(args.Item(i));

        } else {
            // フォルダの場合
            var objFolder = fso.getFolder(args.Item(i) + "\\");
            
            var e = new Enumerator(objFolder.Files);
            for (e.moveFirst(); !e.atEnd(); e.moveNext()) {

                if (boolSupportedFormat(e.item().name)) {
                    
                    WScript.StdOut.WriteLine("Processing : " + e.item().name + " in " + args.Item(i));
                    convertWorkbook2PDF(e.item());
                    
                }
            }
        }
    }

}

function convertWorkbook2PDF (strFileName) {

    var fso = new ActiveXObject("Scripting.FileSystemObject");

    var objExcel = new ActiveXObject("Excel.Application");
    objExcel.visible = false;
    objExcel.displayAlerts = false;

    var wb = objExcel.workbooks.open(strFileName);
    // 上書き
    if (fso.fileExists(wb.path + "\\" + fso.getBaseName(wb.name) + ".pdf") == true) {
        fso.getFile(wb.path + "\\" + fso.getBaseName(wb.name) + ".pdf").Delete(true);

        while (fso.fileExists(wb.path + "\\" + fso.getBaseName(wb.name) + ".pdf") == true) {
            WScript.Sleep(1);
        }
    }
    wb.worksheets.select();
    wb.exportAsFixedFormat(0, wb.path + "\\" + fso.getBaseName(wb.name) + ".pdf");
    wb.close();

    objExcel.quit();
}

main();