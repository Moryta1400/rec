#!/bin/bash
curl -L https://johnvansickle.com/ffmpeg/releases/ffmpeg-release-amd64-static.tar.xz | tar xJ
mkdir -p ~/bin
mv ffmpeg-*/ffmpeg ~/bin/
mv ffmpeg-*/ffprobe ~/bin/
rm -rf ffmpeg-*
