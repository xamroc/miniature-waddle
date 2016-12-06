'use strict'

// const path = require('path')
const nconf = require('nconf')

// Setup nconf to use (in-order):
//   1. Command-line arguments
//   2. Environment variables
nconf.argv().env()

module.exports = nconf.defaults({

  root: __dirname

})
