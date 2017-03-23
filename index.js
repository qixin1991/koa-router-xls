let json2xls = require('json-to-excel'),
    fs = require('fs'),
    xls2json = require('my-xls-to-json');

const date2String = (date, pattern) => {
    if (!date || date == "") return "";
    pattern = pattern || "yyyy-mm-dd HH:MM";
    let str = "";
    str = pattern.replace("yyyy", date.getFullYear());
    str = str.replace("mm", date.getMonth() < 9 ? '0' + (date.getMonth() + 1) : date.getMonth() + 1);
    str = str.replace("dd", date.getDate() < 10 ? '0' + date.getDate() : date.getDate());
    str = str.replace("HH", date.getHours() < 10 ? '0' + date.getHours() : date.getHours());
    str = str.replace("MM", date.getMinutes() < 10 ? '0' + date.getMinutes() : date.getMinutes());
    return str;
}

module.exports = () => {
    return async (ctx, next) => {
        ctx.downloadXLS = (jsonData, fileName) => {
            let xls = json2xls(jsonData);
            ctx.set('Content-Type', 'application/vnd.openxmlformats');
            let filename = encodeURIComponent((fileName || 'download ') + date2String(new Date(), 'yyyy-mm-dd HH-MM') + ".xlsx");
            ctx.set("Content-Disposition", "attachment; filename=" + filename);
            let buffer = new Buffer(xls, 'binary');
            ctx.body = buffer;
        };

        ctx.xlsToJson = (filePath) => {
            let data;
            xls2json(filePath, (err, output) => {
                data = output[Object.keys(output)[0]];   // convert the first worksheet. others do'nt read.(if you need, fork this middleware from github and do what you want.)
                fs.unlink(filePath, (err) => { }); // delete the tmp file.
            });
            return data;
        }
        await next();
    }
}