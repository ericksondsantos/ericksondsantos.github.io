import express from "express";

const app = express();
const port = 3000;

app.get("/", (req, res) => {
  const today = new Date();
  const day = today.getDay();
  var week = "a weekday";
  var time = "it's time to work hard";
  if (day == 0 || day == 6) {
    week = "the weekend";
    time = "it's time to have some fun";
  }
  res.render("/views/index.ejs", {
    dayType: week,
    advice: time,
  });
});

app.listen(port, () => {
  console.log(`Server running on port ${port}.`);
});