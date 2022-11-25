import express from 'express';
const app: express.Express = express();
const port: number = 3000;
app.get('/', (req, res) => {
    res.send('hello~~~')
});
app.listen(port, () => {
    console.log(`listening at http://localhost:${port}`);
});
