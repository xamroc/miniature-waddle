'use strict'

const app = require('koa')()
const router = require('koa-router')()
const path = require('path')

const apiDetails = require('./controllers/api-details')
const apiEvaluate = require('./controllers/api-evaluate')

const zmq = require('zmq')

var reqSock = zmq.socket('req')
reqSock.connect('tcp://127.0.0.1:5556')

router
  .get('/api/details', apiDetails)
  .post('/api/evaluate', apiEvaluate)

app
  .use(router.routes())
  .use(router.allowedMethods())

let port = process.env.PORT || 3000

app.listen(port, () => {
  console.log('Server listening on port ' + port)
})

process.on('SIGINT', function() {
  reqSock.close();
});
