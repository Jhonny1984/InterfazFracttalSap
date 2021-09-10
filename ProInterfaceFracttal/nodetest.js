var request = require("request");
id = 'TLFmgX1Kuef4rsaNxk9z',
    key = '0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU';


var options = {
    method: 'GET',
    url: 'https://app.fracttal.com/api/work_orders_movements/',
    headers:
        {},

    hawk: {
        credentials: {
            id: id,
            key: key,
            algorithm: 'sha256'
        }
    }
};
request(options, function (error, response, body) {
    if (error) throw new Error(error);

    console.log(body);
});