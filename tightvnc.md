# TightVNC

[Download link](https://tightvnc.com/download.php)

## Server

Go through installer (next.. next..)

Let's assume you have two screens configured like this:

```
+---------++---+
|         || 2 |
|    1    |+---+
|         |
+---------+
```

where 1 is main Full HD monitor and 2 is 1024x768 projector. If you want to serve only second monitor on extra port:

Click on TightVNC tray icon. Go to "Extra Ports" tab and click "Add...". In "geometry specification" paste `1024x768+1920+0` and specify port. Hit "OK".

Connect client on a specified port.