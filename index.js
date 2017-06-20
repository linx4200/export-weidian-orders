const rp = require('request-promise-native');
const superagent = require('superagent');
const cheerio = require('cheerio');
const xlsx = require('xlsx');

const url = {
  'login': 'https://sso.weidian.com/user/login',
  'orders': 'https://gwh5.api.weidian.com/wd/order/buyer/getOrderListExt',
  'item': 'https://weidian.com/item.html'
};

const argv = require('minimist')(process.argv.slice(2));
const password = argv.p;
const phone = argv.u;
const outputFile = argv.o || 'output';

// 1. login
async function login (phone, password) {
  const browserMsg = {
    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 9_1 like Mac OS X) AppleWebKit/601.1.46 (KHTML, like Gecko) Version/9.0 Mobile/13B143 Safari/601.1',
    'Content-Type':'application/x-www-form-urlencoded'
  };

  const loginMsg = {
    'countryCode': 86,
    phone,
    password,
    'version': 1
  };

  const response = await superagent.post(url.login).set(browserMsg).send(loginMsg).redirects(0);
  const cookie = response.headers['set-cookie'];

  if (!cookie) {
    console.error('===== 登陆失败 =====');
    return null;
  }

  const res = {};
  cookie.forEach(c => {
    const id =  c.match(/WD_client_userid_raw=(\d+)/);
    if (id) res.id = id[1];

    const wduss =  c.match(/WD_b_wduss=([\d\w]+);/);
    if (wduss) res.wduss = wduss[1];
  });
  
  return res;
}

// 2. fetchData
function parseBody(body) {
  return !body ? {} : JSON.parse(body.replace('jsonp2(', '').replace(/\);?\s*$/g, ''));
}

function writeXlsx(_data) {
  // https://aotu.io/notes/2016/04/07/node-excel/
  const _headers = ['订单号', '商品名', '颜色分类', '数量', '价格'];
  const headers = _headers
                  .map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 }))
                  .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
  const data = _data
                .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65+j) + (i+2) })))
                .reduce((prev, next) => prev.concat(next))
                .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
  // 合并 headers 和 data
  const output = Object.assign({}, headers, data);
  // 获取所有单元格的位置
  const outputPos = Object.keys(output);
  // 计算出范围
  const ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
  // 构建 workbook 对象
  const wb = {
    SheetNames: ['mySheet'],
    Sheets: {
      'mySheet': Object.assign({}, output, { '!ref': ref })
    }
  };
  xlsx.writeFile(wb, `${outputFile}.xlsx`);

  console.info(`===== 写入 ${outputFile}.xlsx 成功 =====`);
}

async function fetchData(param) {
  const res = [];
  const api = `${url.orders}?noticeIsBuyer=1&param=${JSON.stringify(param)}&_=${+new Date()}&callback=jsonp2`;  
  const repos = await rp(api);
  const result = parseBody(repos).result;
  if (result.length) {
    console.info('===== 获取订单信息成功 =====');
    result.forEach(order => {
      order && order.items.forEach(item => {
        res.push({
          'id': item.item_id,
          '订单号': item.order_id,
          // '商品名': title,
          '颜色分类': item.item_sku_title,
          '数量': item.quantity,
          '价格': item.price
        });
      });
    });
  }

  if (res.length) {
    // 需要获取完整的商品名字
    for(let i = 0, l = res.length; i < l; i++){
      const body = await rp(`${url.item}?itemID=${res[i].id}`);
      const $ = cheerio.load(body);
      res[i]['商品名'] = $('head title').text();
    }
    writeXlsx(res.reverse());
  }
}

// run
(async () => {
  if (!password || !phone) {
    console.error('请输入用户名和密码');
    return;
  }

  const info = await login(phone, password);
  if (info) {
    console.info('===== 登陆成功 =====');
  } else {
    return;
  }

  const param = {
    'pageNum': 0,
    'pageSize': 50,
    'ordertype': 'pend',
    'type': 0,
    'userID': info.id,
    'wduss': info.wduss
  };
  fetchData(param);
})();