#!/usr/bin/env bash
workspace=$(cd $(dirname $0) && pwd -P)

{
  cd $workspace
  echo "start cvt doc to docx..."
  osascript doc2docx.scpt
  [[ $? != 0 ]] && exit
  echo "start cvt docx to xlsx one by one..."
  python3 docx2xlsxOneByOne.py
}
