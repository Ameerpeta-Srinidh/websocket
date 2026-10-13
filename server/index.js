import express from "express";
import {WebSocket} from 'ws';

const app = express();
const server = app.listen(8080,()=>{
    console.log("server is running on port 8080");
});

const wss = new WebSocket.Server({server});

wss.on("connection",(ws)=>{
    ws.on("message",(data)=>{
        console.log("message received: ", data);
        ws.send("thank you bud");
    });
});