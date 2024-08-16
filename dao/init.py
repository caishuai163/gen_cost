import sqlite3

from config.excel_conf import DB_URL

init_table = [
    """CREATE TABLE  if not exists ab_tb (
       code INTEGER PRIMARY KEY AUTOINCREMENT,
       user_id INTEGER,
       month_val TEXT,
       rmb NUMBER(10,2),
       is_after INTEGER
   );
   """,
    """CREATE TABLE  if not exists c7_tb (
	code INTEGER PRIMARY KEY AUTOINCREMENT,
	user_id INTEGER,
	month_val TEXT,
	rmb NUMBER(10,2)
);
""",
    """CREATE TABLE  if not exists bound_tb (
	code INTEGER PRIMARY KEY AUTOINCREMENT,
	user_id INTEGER,
	title_id INTEGER,
	rmb NUMBER(10,2));""",
    """CREATE TABLE  if not exists cost_detail_tb (
	code INTEGER PRIMARY KEY AUTOINCREMENT,
	cost_code INTEGER,
	month_val TEXT,
	rmb NUMBER(10,2)
);""", """-- cost_tb definition

CREATE TABLE if not exists  cost_tb (
	code INTEGER PRIMARY KEY AUTOINCREMENT,
		user_id INTEGER,
	file_name TEXT,
	confirm_lost INTEGER,
	avg_cost NUMBER,
	cost_month INTEGER,
	month_cost NUMBER,
	find_count INTEGER,
	confirm_after_find INTEGER
);""", """-- date_tb definition

CREATE TABLE if not exists  date_tb (
	code INTEGER PRIMARY KEY AUTOINCREMENT,
	date_str TEXT
, date REAL);""", """-- p2p_tb definition

CREATE TABLE if not exists  p2p_tb (
	code INTEGER PRIMARY KEY AUTOINCREMENT,
	user_id INTEGER,
	month_val TEXT,
	give_rmb NUMBER(10,2),
	cost_rmb NUMBER(10,2),
	cost_month TEXT
);""", """-- spec_tb definition

CREATE TABLE if not exists  spec_tb (
	code INTEGER PRIMARY KEY AUTOINCREMENT,
	user_id INTEGER,
	month_val TEXT,
	cost_rmb NUMBER(10,2),
	cost_month TEXT
);""", """-- user_tb definition

CREATE TABLE if not exists  user_tb (
	code INTEGER PRIMARY KEY AUTOINCREMENT,
	user_name TEXT,
	region TEXT,
	region_manager TEXT,
	part TEXT,
	company TEXT,	
	"work" TEXT, 
	dealer_name TEXT,
	platform TEXT, need_generated INTEGER);"""
]


def clear_db():
    print("clear old data")
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    execute.execute("delete from bound_tb;")
    execute.execute("delete from cost_tb;")
    execute.execute("delete from cost_detail_tb;")

    execute.execute("delete from p2p_tb;")
    execute.execute("delete from ab_tb;")
    conn.commit()
    execute.close()
    print("clear old data success")


def init():
    conn = sqlite3.connect(DB_URL)
    execute = conn.cursor()
    for sql in init_table:
        execute.execute(sql)
    conn.commit()
    execute.close()
