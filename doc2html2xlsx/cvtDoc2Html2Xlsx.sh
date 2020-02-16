#!/usr/bin/env bash
workspace=$(cd $(dirname $0) && pwd -P)

{
  cd $workspace
  echo "start cvt doc to html..."
  osascript doc2html.scpt
  [[ $? != 0 ]] && exit
  echo "start cvt html to xlsx combined..."
  python3 html2xlsxCombined.py
}
