'use strict'

const _ = require('lodash')
const formidable = require('koa-formidable')
const DataTransform = require("node-json-transform").DataTransform
const zmq = require('zmq')

var reqSock = zmq.socket('req')
reqSock.connect('tcp://127.0.0.1:5556')

module.exports = function * () {
  const form = yield formidable.parse(this)

  var inputOverride = { input: form.fields }

  try {
    let results = yield new Promise(resolve => {
      reqSock.on('message', function(msg) { resolve(JSON.parse(msg)) })
      reqSock.send(JSON.stringify(inputOverride))
    })

    this.body = results
  } catch (err) {
    this.throw(err, 500)
  }
}

process.on('SIGINT', function() {
  reqSock.close();
});