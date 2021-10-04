#!/bin/bash
python prep_data.py
pptx_chart -o presentation1.pptx -d chart1.csv
pptx_chart -o presentation2.pptx -d chart2.csv
