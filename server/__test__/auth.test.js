process.env.NODE_ENV = 'test';

const request = require('supertest');
const app = require('../app');
const db = require('../database/db');

describe('Connection to database succesful', () => {
  test('can query the database', async () => {
    const result = await db.query('SELECT 1 + 1 AS result');
    expect(result.rows[0].result).toBe(2);
  });
});

describe('Create a new user (sign up)', () => {
  test('create a user', async () => {
    const res = await request(app).post('/auth/signup').send({
      email: 'testemail@email.com',
      password: 'password123',
      first_name: 'John',
      last_name: 'Doe',
    });
    expect(res.body.message).toBeDefined();
    expect(res.body.user).toBeDefined();
  });

  test('Password too short', async () => {
    const res = await request(app)
      .post('/auth/signup')
      .send({ email: 'testemail@email.com', password: 'p1' });
    expect(res.body.error).toBe('Validation Error');
    expect(res.body.errors.length).toBe(1);
    expect(res.body.errors[0].includes('minimum length of 8')).toBe(true);
  });

  test('Password missing number and/or letter', async () => {
    const res = await request(app)
      .post('/auth/signup')
      .send({ email: 'testemail@email.com', password: '123456789' });
    expect(res.body.error).toBe('Validation Error');
    expect(res.body.errors[0].includes('does not match allOf schema')).toBe(
      true
    );
  });

  test('Email not an email', async () => {
    const res = await request(app)
      .post('/auth/signup')
      .send({ email: 'testemailemail.com', password: 'pasword123' });
    expect(res.body.error).toBe('Validation Error');
    expect(
      res.body.errors[0].includes('does not conform to the "email" format')
    ).toBe(true);
  });
});

describe('Log in', () => {
  test('Log in with correct credentials', async () => {
    const res = await request(app)
      .post('/auth/login')
      .send({ email: 'testemail@email.com', password: 'password123' });
    expect(res.body.message).toBe('Successful login');
    expect(res.body.user).toBeDefined();
  });

  test('Log in with incorrect password', async () => {
    const res = await request(app)
      .post('/auth/login')
      .send({ email: 'testemail@email.com', password: 'wrongPASWORD123' });
    expect(res.body.message).toBe('Inputted password is incorrect');
  });

  test('Log in with non-existant email', async () => {
    const res = await request(app)
      .post('/auth/login')
      .send({ email: 'testemailWRONG@email.com', password: 'password123' });
    expect(res.body.message).toBe(
      'Inputted email (testemailWRONG@email.com) is not in the userbase'
    );
  });
});

describe('Log out', () => {
  test('Test logging out', async () => {
    const res = await request(app).get('/auth/logout');
    expect(res.body.message).toBe('Successful logout');
    expect(res.body.access).toBe(false);
  });
});

afterAll(async function () {
  await db.query('DELETE FROM users');
  await db.end();
});
