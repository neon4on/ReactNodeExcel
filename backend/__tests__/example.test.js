const request = require('supertest');
const app = require('../Main'); // Подключаем ваш файл Main.js

test('Server is listening at specified port', async () => {
  const response = await request(app).get('/');
  expect(response.statusCode).toBe(200);
  expect(response.text).toContain('Server is listening');
});

afterAll((done) => {
  server.close(done);
});
