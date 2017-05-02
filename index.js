const rp = require('request-promise-native');
const cheerio = require('cheerio');
var xlsx = require('xlsx');

function parseBody(body) {
  return !body ? {} : JSON.parse(body.replace('jsonp2(', '').replace(/\);?\s*$/g, ''));
}

const api = 'https://gwh5.api.weidian.com/wd/order/buyer/getOrderListExt';

const param = {
  'pageNum': 0,
  'pageSize': 50,
  'ordertype': 'pend',
  'type': 2,
  'userID': '',
  'wduss': ''
};

const url = `${api}?noticeIsBuyer=1&param=${JSON.stringify(param)}&_=${+new Date()}&callback=jsonp2`;

const res = [];

(async () => {
  const repos = await rp(url);
  const result = parseBody(repos).result;
  if (result.length) {
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
    for(let i = 0, l = res.length; i < l; i++){
      const body = await rp(`https://weidian.com/item.html?itemID=${res[i].id}`);
      const $ = cheerio.load(body)
      res[i]['商品名'] = $('head title').text()
    }
    console.log(res)

    var _headers = ['订单号', '商品名', '颜色分类', '数量', '价格']
    var _data = res.reverse();
    var headers = _headers
                    .map((v, i) => Object.assign({}, {v: v, position: String.fromCharCode(65+i) + 1 }))
                    .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
    var data = _data
                  .map((v, i) => _headers.map((k, j) => Object.assign({}, { v: v[k], position: String.fromCharCode(65+j) + (i+2) })))
                  .reduce((prev, next) => prev.concat(next))
                  .reduce((prev, next) => Object.assign({}, prev, {[next.position]: {v: next.v}}), {});
    // 合并 headers 和 data
    var output = Object.assign({}, headers, data);
    // 获取所有单元格的位置
    var outputPos = Object.keys(output);
    // 计算出范围
    var ref = outputPos[0] + ':' + outputPos[outputPos.length - 1];
    // 构建 workbook 对象
    var wb = {
        SheetNames: ['mySheet'],
        Sheets: {
            'mySheet': Object.assign({}, output, { '!ref': ref })
        }
    };

    // var wopts = { bookType:'xlsx', bookSST:false, type:'binary' };
    // var wbout = xlsx.write(wb,wopts);
    xlsx.writeFile(wb, 'output.xlsx');
  }
})();