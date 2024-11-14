import { config } from 'dotenv';
import express from 'express';
import cors from 'cors';

config();

const port: string | undefined = process.env.PORT || "3000";

async function main() {
    const app: express.Express = express();
    app.use(cors());
    app.use(express.static("./dist"));
    app.listen(
        port,
        () => console.log(`Server running on port: ${port}`)
    );
}

main();