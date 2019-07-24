from flask_restful import Resource
from flask import request, jsonify
import xlrd
import pymysql
import json
pymysql.install_as_MySQLdb()


class Excel(Resource):
    def post(self):
        resp = request.files.get("excel")
        if resp is None:
            return "Không tìm thấy file", 404
        workbook = xlrd.open_workbook(file_contents=resp.read())
        try:
            worksheet = workbook.sheet_by_name("Sheet1")
        except:
            return "Không tìm thấy sheet", 404
        try:
            db = pymysql.connect('127.0.0.1', 'root', 'thanh1997', 'excel')
            cursor = db.cursor()
            cursor.execute('CREATE TABLE IF NOT EXISTS asset '
                           '(id INT PRIMARY KEY AUTO_INCREMENT,'
                           ' stt VARCHAR(30), '
                           'danh_muc_tai_san VARCHAR(500),'
                           ' thoi_gian_su_dung  VARCHAR(30) ,'
                           ' ti_le_hao_mon VARCHAR(30))');
            list_data = []
            for r in range(2, worksheet.nrows):
                cursor.execute('insert into asset (stt, danh_muc_tai_san, thoi_gian_su_dung, ti_le_hao_mon) values '
                               '("{0}", "{1}", "{2}", "{3}");'
                               .format(worksheet.cell(r, 0).value, worksheet.cell(r, 1).value,
                                       worksheet.cell(r, 2).value, worksheet.cell(r, 3).value))
                if worksheet.cell(r, 2).value == "" and worksheet.cell(r, 3).value == "":
                    danh_muc_tai_san = worksheet.cell(r, 1).value
                    continue
                list_data.append({"danh_muc_tai_san": danh_muc_tai_san,
                                  "thoi_gian_su_dung": worksheet.cell(r, 2).value,
                                  "ti_le_hao_mon": worksheet.cell(r, 3).value,
                                  "loai_tai_san": worksheet.cell(r, 1).value[2:]})
            db.commit()
        except:
            return "Lỗi không xác định", 500
        return {"list data": list_data}, 200
