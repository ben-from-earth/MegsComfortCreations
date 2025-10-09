const express = require('express');
const router = new express.Router();
const jwt = require('jsonwebtoken');
const userCreateSchema = require('../schemas/userCreateSchema.json');
const bcrypt = require('bcrypt');
const User = require('../models/user');
const { validate } = require('jsonschema');

const secret = process.env.JWTSECRET;

const maxAge = 3 * 24 * 60 * 60;
const createToken = (id) => {
  return jwt.sign({ id }, secret, {
    expiresIn: maxAge,
  });
};

const handleErrors = (res, err) => {
  if (err.error.code == 23505) {
    return res.status(400).json({
      error: 'Duplication Error',
      message: 'That email is already registered',
    });
  }
  if (err.error === 'User Creation error') {
    return res.status(400).json({
      error: err.error,
      message:
        'There was a user creation error or the database is currently offline',
    });
  }

  if (err.error === 'Validation Error') {
    return res.status(400).json({
      error: 'Validation Error',
      errors: err.validation.errors.map((e) => e.stack),
    });
  }

  if (err.error === 'Incorrect Password') {
    return res.status(403).json({
      error: 'Authentication Error',
      message: 'Inputted password is incorrect',
    });
  }

  if (err.error === 'Email Incorrect') {
    return res.status(403).json({
      error: 'Authentication Error',
      message: `Inputted email (${err.email}) is not in the userbase`,
    });
  }
};

router.post('/signup', async (req, res) => {
  const { email, password, first_name, last_name } = req.body;
  const salt = await bcrypt.genSalt();
  const hashedpass = await bcrypt.hash(password, salt);
  try {
    const validation = validate(req.body, userCreateSchema);
    if (!validation.valid) {
      return handleErrors(res, { error: 'Validation Error', validation });
    }

    const user = await User.create({
      email,
      password: hashedpass,
      first_name,
      last_name,
    });

    if (user.id) {
      const token = createToken(user.id);
      res.cookie('jwt', token, { httpOnly: true, maxAge: maxAge * 1000 });
      res.status(201).json({
        message: 'Successfully added user',
        user,
      });
    } else {
      return handleErrors(res, { error: 'User Creation Error' });
    }
  } catch (err) {
    console.log(err);
    return handleErrors(res, { error: err });
  }
});
router.post('/login', async (req, res) => {
  const { email, password } = req.body;
  const user = await User.login(email);
  if (user) {
    const auth = await bcrypt.compare(password, user.password_hash);
    const token = createToken(user.id);
    if (auth) {
      res.cookie('jwt', token, {
        httpOnly: true,
        maxAge: maxAge * 1000,
      });
      res.status(200).json({ user: user, message: 'Successful login' });
    } else {
      return handleErrors(res, { error: 'Incorrect Password' });
    }
  } else {
    return handleErrors(res, { error: `Email Incorrect`, email });
  }
});
router.get('/logout', (req, res) => {
  res.clearCookie('jwt');
  res.status(200).json({ access: false, message: 'Successful logout' });
});

router.get('/me', (req, res) => {
  const token = req.cookies?.jwt;
  if (token) {
    jwt.verify(token, secret, async (err, decodedToken) => {
      if (err) {
        res.status(403).json({
          error: 'Access Denied',
          message: 'This route is protected, please sign in or register',
        });
      } else {
        let user = await User.findOne(decodedToken.id);

        res.status(200).json({ access: true, user });
      }
    });
  } else {
    res.status(403).json({
      error: 'Access Denied',
      message: 'This route is protected, please sign in or register',
    });
  }
});

module.exports = router;
