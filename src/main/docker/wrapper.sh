#!/bin/bash

java -Djava.security.egd=file:/dev/./urandom -Dspring.profiles.active=container -Duser.timezone=Asia/Shanghai -jar /app.jar