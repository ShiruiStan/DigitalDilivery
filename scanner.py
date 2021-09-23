# -- coding: utf-8 --
# @Time : 2021/9/23 16:10
# @Author : Shi Rui

import sqlite3
import xml.etree
import os


def analyse_schema(schema):
    for file in os.listdir(schema):
        pass


class Scanner:
    def __init__(self, db, schema, code):
        analyse_schema(schema)

    def start(self):
        pass
