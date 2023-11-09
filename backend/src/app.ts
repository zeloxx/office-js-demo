import express from "express";

const app = express();

app.use(express.json());

app.get("/marco", (req, res) => {
  res.send("polo");
});

export default app;
