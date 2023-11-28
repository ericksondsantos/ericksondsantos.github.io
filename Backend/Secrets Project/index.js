import express from "express";
import bodyParser from "body-parser";
import { dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));

var pass = "ILoveProgramming";
var passwordAccepted = false;

const app = express();
const port = 3000;

app.use(bodyParser.urlencoded({extended: true}));

function checkPassword(req, res, next){
    const formPassword = req.body["password"];
    if(formPassword == pass){
        passwordAccepted = true;
    }
    next();
}

app.use(checkPassword);

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/public/index.html");
});

app.post("/check", (req, res) => {
    if(passwordAccepted == true){
        res.sendFile(__dirname + "/public/secret.html");
    }
    else{
        res.sendFile(__dirname + "/public/index.html");
    }
});

app.listen(port, () => {
  console.log(`Listening on port ${port}`);
});
