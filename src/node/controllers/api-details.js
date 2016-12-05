'use strict'

const process = require('process')

const _ = require('lodash')
const formidable = require('koa-formidable')
const DataTransform = require("node-json-transform").DataTransform
const zmq = require('zmq')

module.exports = function * () {
  
    // let spreadsheet = { id: 'pricer.xlsm', purpose: 'pricer spreadsheet for POC' }

    const details = require('../data/details.json')
    const inputs = require('../data/inputs.json')
    // let inputs = {"input":[{"Name":"Auth_Dis","FgColor":"#3B689F","BgColor":"#C00000","ValueType":"","Value":"","DataType":"Empty","HTMLDataType":"text"},{"Name":"AXA_Premium_Workshop","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"$0.00","DataType":"Currency","HTMLDataType":"text"},{"Name":"Claims_Amount","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"$0.00","DataType":"Currency","HTMLDataType":"text"},{"Name":"Comm_Giveaway","FgColor":"#3B689F","BgColor":"#C00000","ValueType":"","Value":"","DataType":"Empty","HTMLDataType":"text"},{"Name":"Cover","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"Comprehensive","DataType":"Text","HTMLDataType":"text"},{"Name":"Driver1_Occupation","FgColor":"#000000","BgColor":"#EBF1DE","ValueType":"","Value":"Driver","DataType":"Text","HTMLDataType":"text"},{"Name":"Driver2_Occupation","FgColor":"#000000","BgColor":"#C00000","ValueType":"","Value":"Driver","DataType":"Text","HTMLDataType":"text"},{"Name":"Driver3_Occupation","FgColor":"#000000","BgColor":"#EBF1DE","ValueType":"","Value":"Driver","DataType":"Text","HTMLDataType":"text"},{"Name":"Driver4_Occupation","FgColor":"#000000","BgColor":"#EBF1DE","ValueType":"","Value":"Driver","DataType":"Text","HTMLDataType":"text"},{"Name":"Driver5_Occupation","FgColor":"#000000","BgColor":"#EBF1DE","ValueType":"","Value":"Driver","DataType":"Text","HTMLDataType":"text"},{"Name":"ExpiryDate","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"","Value":"2017-12-31","DataType":"Date","HTMLDataType":"date"},{"Name":"First_Inception_Date","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"1900-01-00","DataType":"Date","HTMLDataType":"date"},{"Name":"Inception_Date","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"2017-01-01","DataType":"Date","HTMLDataType":"date"},{"Name":"Last_Year_Basic_Premium","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"$0.00","DataType":"Currency","HTMLDataType":"text"},{"Name":"Last_Year_NCD","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"0%","DataType":"Percent","HTMLDataType":"number"},{"Name":"Last_Year_NCD_Protector","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"0%","DataType":"Percent","HTMLDataType":"number"},{"Name":"Last_Year_Plan_Premium","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"$0.00","DataType":"Currency","HTMLDataType":"text"},{"Name":"Last_Year_SDD","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"0%","DataType":"Percent","HTMLDataType":"number"},{"Name":"Last_Year_Voluntary_Excess","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"$0.00","DataType":"Currency","HTMLDataType":"text"},{"Name":"Main_Driver_DOB","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"1980-01-01","DataType":"Date","HTMLDataType":"date"},{"Name":"Main_Driver_Gender","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"Male","DataType":"Text","HTMLDataType":"text"},{"Name":"Main_Driver_Occupation","FgColor":"#000000","BgColor":"#EBF1DE","ValueType":"","Value":"teacher","DataType":"Text","HTMLDataType":"text"},{"Name":"Nationality","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"1900-01-00","DataType":"Date","HTMLDataType":"date"},{"Name":"NBRN","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"New Business","DataType":"Text","HTMLDataType":"text"},{"Name":"NCD","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"40%","DataType":"Percent","HTMLDataType":"number"},{"Name":"No_of_Named_Drivers","FgColor":"#0000FF","BgColor":"#C00000","ValueType":"","Value":"5","DataType":"Number","HTMLDataType":"number"},{"Name":"Occupation","FgColor":"#000000","BgColor":"#EBF1DE","ValueType":"","Value":"teacher","DataType":"Text","HTMLDataType":"text"},{"Name":"off_peak_car","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"Yes","DataType":"Text","HTMLDataType":"text"},{"Name":"Postcode","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"534122","DataType":"Number","HTMLDataType":"number"},{"Name":"sel_and1dob","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"","Value":"1982-12-03","DataType":"Date","HTMLDataType":"date"},{"Name":"sel_and1drivexp","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"4 - 5 years","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and1gender","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Male","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and1marital","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Single, Divorced, or Widowed","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and1relation","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Parent","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and2dob","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"","Value":"1982-12-04","DataType":"Date","HTMLDataType":"date"},{"Name":"sel_and2drivexp","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"4 - 5 years","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and2gender","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Male","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and2marital","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Single, Divorced, or Widowed","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and2relation","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Parent","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and3dob","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"","Value":"1982-12-04","DataType":"Date","HTMLDataType":"date"},{"Name":"sel_and3drivexp","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"4 - 5 years","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and3gender","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Male","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and3marital","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Single, Divorced, or Widowed","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and3relation","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Parent","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and4dob","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"","Value":"1982-12-04","DataType":"Date","HTMLDataType":"date"},{"Name":"sel_and4drivexp","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"4 - 5 years","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and4gender","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Male","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and4marital","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Single, Divorced, or Widowed","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and4relation","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Parent","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and5dob","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"","Value":"1982-12-04","DataType":"Date","HTMLDataType":"date"},{"Name":"sel_and5drivexp","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"4 - 5 years","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and5gender","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Male","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and5marital","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Single, Divorced, or Widowed","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_and5relation","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Parent","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_interm","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"03180 - KHC HOLDINGS PTE LTD","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_lycaraccessoriesamount","FgColor":"#3B689F","BgColor":"#C00000","ValueType":"","Value":"","DataType":"Empty","HTMLDataType":"text"},{"Name":"sel_manualCR1b","FgColor":"#FFFFFF","BgColor":"#C00000","ValueType":"","Value":"$0.00","DataType":"Currency","HTMLDataType":"text"},{"Name":"sel_manualDA1b","FgColor":"#FFFFFF","BgColor":"#C00000","ValueType":"","Value":"$0.00","DataType":"Currency","HTMLDataType":"text"},{"Name":"sel_maritail","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"List","Value":"Married","DataType":"Text","HTMLDataType":"text"},{"Name":"sel_quodate","FgColor":"#3B689F","BgColor":"#EBF1DE","ValueType":"","Value":"2016-09-29","DataType":"Date","HTMLDataType":"date"},{"Name":"sel_renwsclaim","FgColor":"#3B689F","BgColor":"#C00000","ValueType":"","Value":"0","DataType":"Number","HTMLDataType":"number"},{"Name":"sel_rewsclaim","FgColor":"#3B689F","BgColor":"#C00000","ValueType":"","Value":"0","DataType":"Number","HTMLDataType":"number"},{"Name":"Vehicle_Make","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"TALBOT","DataType":"Text","HTMLDataType":"text"},{"Name":"Vehicle_Model","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"MATR 1.4","DataType":"Text","HTMLDataType":"text"},{"Name":"Vehicle_Yr_Manufacture","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"$2,012.00","DataType":"Currency","HTMLDataType":"text"},{"Name":"Year_Driving_License_obtained","FgColor":"#0000FF","BgColor":"#EBF1DE","ValueType":"","Value":"2 - 3 years","DataType":"Text","HTMLDataType":"text"}]}

    this.body = { id: details.id, purpose: details.description, inputs: inputs }
}