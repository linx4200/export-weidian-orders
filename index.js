const rp = require('request-promise');
// const cheerio = require('cheerio');
const tough = require('tough-cookie');
const Cookie = tough.Cookie;
const xlsx = require('xlsx');

const url = {
  'login': 'https://sso.weidian.com/user/login',
  'orders': 'https://gwh5.api.weidian.com/wd/buyer/buyer_query_order_list',
  'item': 'https://weidian.com/item.html'
};

const argv = require('minimist')(process.argv.slice(2));
const password = argv.p;
const phone = argv.u;
const outputFile = argv.o || 'output';

function request(options) {
  return new Promise((resolve, reject) => {
    function autoParse(body, response, resolveWithFullResponse) {
      return [body, response, resolveWithFullResponse];
    }
    options.transform = autoParse;
    rp(options).then(body => resolve(body)).catch(err => reject(err));
  });
}

// 1. login
async function login (phone, password) {

  const form = {
    'countryCode': 86,
    phone,
    password,
    'version': 1
  };

  const options = {
    method: 'POST',
    headers: {
      referer: 'https://sso.weidian.com/login/index.php'
    },
    uri: url.login,
    form,
    json: true // Automatically stringifies the body to JSON
  };

  const [body, response] =  await request(options);
  
  if (body.status && body.status.status_code !== 0) {
    return new Error('===== 登陆失败 =====');
  }
  let cookies;
  if (response.headers['set-cookie'] instanceof Array) {
    cookies = response.headers['set-cookie'].map(Cookie.parse);
  } else {
    cookies = [Cookie.parse(response.headers['set-cookie'])];
  }

  const res = {};
  res.cookies = cookies;
  cookies.forEach(c => {
    if (c.key === 'uid') res.id = c.value;
    if (c.key === 'WD_b_wduss') res.wduss = c.value;
  });
  return res;
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

// 2. fetchData
async function fetchData(param) {
  const res = [];

  const cookies = param.cookies;
  const cookiejar = rp.jar();
  const domain = 'https://gwh5.api.weidian.com';
  cookies.forEach(cookie => cookiejar.setCookie(cookie, domain));
  delete param.cookies;

  const options = {
    uri: url.orders,
    qs: {
      noticeIsBuyer: '1',
      param: JSON.stringify(param),
      _: +new Date()
    },
    headers: {
      referer: 'https://i.weidian.com/order/list.php?type=0'
    },
    jar: cookiejar,
    json: true // Automatically parses the JSON string in the response
  };

  let body;
  try {
    [body] = await request(options);
  } catch (e) {
    new Error(`===== 获取数据失败 ===== ${e.message}`);
  }

  if (body.result && body.result.length) {
    console.info('===== 获取订单信息成功 =====');
    body.result.forEach(order => {
      order && order.items.forEach(item => {
        res.push({
          'id': item.item_id,
          '订单号': order.order_id,
          '商品名': item.item_title,
          '颜色分类': item.item_sku_title,
          '数量': item.quantity,
          '价格': item.price
        });
      });
    });
  }

  if (res.length) {
    // 需要获取完整的商品名字
    // for(let i = 0, l = res.length; i < l; i++){
    //   const body = await rp(`${url.item}?itemID=${res[i].id}`);
    //   const $ = cheerio.load(body);
    //   res[i]['商品名'] = $('head title').text();
    // }
    writeXlsx(res.reverse());
  }
}

// run
(async () => {
  if (!password || !phone) {
    console.error('请输入用户名和密码');
    return;
  }
  let info;
  try {
    info = await login(phone, password);
  } catch (e) {
    console.error(`==== 登陆失败 ==== ${e.message}`);
    return;
  }

  console.info('===== 登陆成功 =====');

  const param = {
    page: 0,
    page_size: 100,
    type: 0,
    buyer_id: info.id,
    wduss: info.wduss,
    cookies: info.cookies
  };
  try {
    fetchData(param);
  } catch (e) {
    console.error(e);
    return;
  }
})();