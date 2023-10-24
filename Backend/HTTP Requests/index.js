import express from "express";

const app = express();
const port = 3000;

app.get("/", (req, res) => {
    res.send("Hello World!");
});

// above syntax can also be written as:
// app.get("/", sendMessage(req, res));
// function sendMessage(req, res){
//     res.send("Hello World!");
// }

app.listen(port, () => {
    console.log(`Server is running on port ${port}.`);
});

// above syntax can also be written as:
// app.listen(port, consoleMessage())
// function consoleMessage(){
//     console.log(`Server is runnint on port ${port}.`);
// }

