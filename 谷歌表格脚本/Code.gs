/**
 * 扫码点单系统 - Google Apps Script 后端
 *
 * 与 客户端/customer.html、员工端/sjxddxtygd.html 配套使用，
 * 不需要任何自建服务器，数据全部存在这张 Google 表格里。
 *
 * 部署步骤（一次性，大约 3 分钟）：
 * 1. 打开 https://sheet.new 新建一个空白 Google 表格（名字随意）
 * 2. 顶部菜单：扩展程序 → Apps Script，会打开脚本编辑器
 * 3. 把默认的 Code.gs 内容全部删掉，粘贴本文件的全部内容，Ctrl+S 保存
 * 4. 右上角蓝色按钮"部署" → "新建部署"
 *    - 点类型旁的齿轮图标，选择"Web 应用"
 *    - 执行身份：选"我"（你自己的账号）
 *    - 有权访问的用户：选"任何人"
 *    - 点击"部署"
 * 5. 首次部署会弹出 Google 授权提示，按提示一路允许（不受信任脚本→仍要前往→允许）
 * 6. 部署完成后复制那个以 /exec 结尾的网址
 * 7. 把这个网址发给我，我来帮你替换进两个 HTML 文件的 API_URL 里
 *    （或者你自己手动替换 customer.html 和 sjxddxtygd.html 里的 API_URL 常量）
 *
 * 以后如果改了这份脚本代码，需要重新"管理部署"→选中已有部署→编辑→
 * 新版本→部署，网址不会变，不用改前端。
 */

const SHEET_NAME = 'Orders';
const HEADERS = ['id', 'table', 'items', 'createdAt', 'status'];

function getSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if(!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function doGet(e){
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const rows = data.slice(1); // 去掉表头行
  const orders = rows
    .filter(r => r[0] && r[4] !== 'done')
    .map(r => ({
      id: r[0],
      table: r[1],
      items: JSON.parse(r[2]),
      createdAt: Number(r[3]),
    }));
  return jsonOutput(orders);
}

function doPost(e){
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet();

  if(body.action === 'create'){
    const id = Utilities.getUuid();
    sheet.appendRow([id, String(body.table), JSON.stringify(body.items), body.createdAt, 'pending']);
    return jsonOutput({ ok: true, id: id });
  }

  if(body.action === 'complete'){
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++){
      if(data[i][0] === body.id){
        sheet.getRange(i + 1, 5).setValue('done'); // 第5列是 status
        break;
      }
    }
    return jsonOutput({ ok: true });
  }

  return jsonOutput({ ok: false, error: 'unknown action: ' + body.action });
}

function jsonOutput(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
