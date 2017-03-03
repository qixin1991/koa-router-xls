# koa-router-xls

It's a middleware of `koa-router@next` for xls and json convert to each other.

## Install

```
npm install koa-router-xls
```

## Usage

**app.js**

```
"use strict";
const Koa = require('koa'),
      app = new Koa(),
      xls = require('koa-router-xls'),
      formParser = require('koa-router-form-parser'),
      userRouter = require('./routes/user.js');

app.use(formParser());
app.use(xls());

app.use(userRouter.routes());

var port = 3000;
app.listen(port, function () {
  console.log(` ---> Server running on port: ${port}`);
});
```

**routes/user.js**

```
"use strict";
router.get('/download', async ctx => {
    var data = [{ name: "brain", age: 24 }, { name: "无忌", age: 26 }];
    ctx.downloadXLS(data,'mydownload-xls');
});

// NOTE: this method dependency on `koa-router-form-parse` and 
//  it'll be automatic installed when you install `koa-router-xls`
router.post('/uploadxls', async ctx => {
    var filePath = await ctx.formParse(); // ctx.formParse() is a extend middleware named `koa-router-form-parser` of koa-router@next
    var data = ctx.xlsToJson(filePath);
    ctx.body = data;
});
```
