import { config } from 'dotenv';
import express from 'express';

config();

const port: string | undefined = process.env.PORT || "3000";

async function main() {
    const app: express.Express = express();
    app.use(express.static("./dist"));
    app.listen(
        port,
        () => console.log(`Server running on port: ${port}`)
    );
}

main();