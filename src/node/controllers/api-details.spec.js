'use strict';
const chai = require('chai');
const expect = chai.expect;
const should = chai.should();
const sinon = require('sinon');
const Promise = require('bluebird')
const details = require('./api-details')

let app = require('../index.js')
let request = require('supertest').agent(app.listen())

describe('#details', () => {

  it('should return inputs', (done) => {
    request
      .get('/api/details')
      .expect(200)
      .end((err, res) => {
        if (err) return done(err);
        done();
      })
  })
})
