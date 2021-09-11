const express= require('express');
const session = require('express-session');
const cookieParser = require('cookie-parser');
const app= express();
const path= require('path');
const msal = require('@azure/msal-node');
const bodyParser = require('body-parser');

var hbs = require('hbs');
require('dotenv').config();
const flash= require('connect-flash');

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(cookieParser());
app.use(session({
    secret: 'demo_secret',
    resave: true,
    saveUninitialized: false,
    unset: 'destroy'
  }));
  
app.locals.users = {};

const msalConfig = {
    auth: {
      clientId: process.env.OAUTH_CLIENT_ID,
      authority: process.env.OAUTH_AUTHORITY,
      clientSecret: process.env.OAUTH_CLIENT_SECRET
    },
    system: {
      loggerOptions: {
        loggerCallback(loglevel, message, containsPii) {
          console.log(message);
        },
        piiLoggingEnabled: false,
        logLevel: msal.LogLevel.Verbose,
      }
    }
  };

app.locals.msalClient = new msal.ConfidentialClientApplication(msalConfig);

app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'hbs');
app.use(express.static(path.join(__dirname, 'public')));
app.use(flash());

app.use((req, res, next)=> {
    // Read any flashed errors and save
    // in the response locals
    res.locals.error = req.flash('error_msg');
  
    // Check for simple error string and
    // convert to layout's expected format
    var errs = req.flash('error');
    for (var i in errs){
      res.locals.error.push({message: 'An error occurred', debug: errs[i]});
    }
  
    // Check for an authenticated user and load
    // into response locals
    if (req.session.userId) {
      res.locals.user = app.locals.users[req.session.userId];
    }
  
    next();
  });

app.get('/', (req, res) => {
    res.render('index.hbs', {active: { home: true }});
})

const authRouter= require('./auth')
const mail= require('./mail');

app.use('/auth', authRouter);
app.use('/mail', mail);

app.listen(3000);