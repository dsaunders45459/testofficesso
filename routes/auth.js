

var express = require('express');
var jsonwebtoken = require('jsonwebtoken');
var router = express.Router();
var fetch = require('node-fetch');
var form = require('form-urlencoded').default;


/* GET users listing. */
router.get('/', async function(req, res, next) {
  const authorization = req.get('Authorization');
  if (authorization == null) {
      throw new Error('No Authorization header was found.');
  }
  const [schema, jwt] = authorization.split(' ');

  const decoded = jsonwebtoken.decode(jwt, { complete: true });
  const v2Params = {
    client_id: process.env.CLIENT_ID,
    client_secret: process.env.CLIENT_SECRET,
    grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
    assertion: jwt,
    requested_token_use: 'on_behalf_of',
    scope: ['user.read'].join(' ')
  };

  const stsDomain = 'https://login.microsoftonline.com';
  const tenant = 'common';
  const tokenURLSegment = 'oauth2/v2.0/token';

  const tokenResponse = await fetch(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
    method: 'POST',
    body: form(v2Params),
    headers: {
        'Accept': 'application/json',
        'Content-Type': 'application/x-www-form-urlencoded'
    }
  });
  const json = await tokenResponse.json();
  
  res.send(json);
});

router.get('/getuserdata', async function(req, res, next) {
  const authorization = req.get('access_token');

  const tokenResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages`, {
    method: 'GET',
    headers: {
      "Authorization": `Bearer ${authorization}`
    }
  });
  const json = await tokenResponse.json();
  
  res.send(json);
});

router.get('/token', async function(req, res, next) {
  const authorization = req.query.code;
  
  res.send(authorization);
});
module.exports = router;
