import express from "express";

const app = express();
const port = 3000;

app.get("/", (req, res) => {
    const today = new Date();
    const day = today.getDay();
    if (day === 0 || day === 6){
        res.render("index.ejs", {dayType: "weekend", advise: "it's time to play hard"});
    }
    else{
        res.render("index.ejs", {dayType: "weekday", advise: "it's time to work hard"});
    }
});

app.listen(port, () => {
    console.log(`Listening on port ${port}`);
});
