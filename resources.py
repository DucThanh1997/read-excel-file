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
            return {"messages": "Không tìm thấy file"}, 404
        workbook = xlrd.open_workbook(file_contents=resp.read())
        try:
            worksheet = workbook.sheet_by_name("Sheet1")
        except:
            return {"messages": "Không tìm thấy sheet"}, 404
        try:
            db = pymysql.connect("127.0.0.1", "root", "thanh1997", "excel")
            cursor = db.cursor()
            cursor.execute(
                "CREATE TABLE IF NOT EXISTS asset "
                "(id INT PRIMARY KEY AUTO_INCREMENT,"
                " stt VARCHAR(30), "
                "danh_muc_tai_san VARCHAR(500),"
                " thoi_gian_su_dung  VARCHAR(30) ,"
                " ti_le_hao_mon VARCHAR(30))"
            )
            response_data = []
            for r in range(2, worksheet.nrows):
                cursor.execute(
                    "insert into asset (stt, danh_muc_tai_san, thoi_gian_su_dung, ti_le_hao_mon) values "
                    '("{0}", "{1}", "{2}", "{3}");'.format(
                        worksheet.cell(r, 0).value,
                        worksheet.cell(r, 1).value,
                        worksheet.cell(r, 2).value,
                        worksheet.cell(r, 3).value,
                    )
                )
                
                # lấy các giá trị danh mục lớn kiểm tra xem nó có chữ loại ở đầu không và 2 cell sau có trống không
                if (
                    str(worksheet.cell(r, 0).value).find("Loại") != -1
                    and worksheet.cell(r, 2).value == ""
                    and worksheet.cell(r, 3).value == ""
                ):
                    danh_muc_tai_san_lon = worksheet.cell(r, 1).value
                    is_subtitle = 0
                    continue
                # lấy giá trị danh mục con, nó ko có chữ loại ở đầu và nó cũng ko trống và 2 cell sau của nó có trống
                # không?
                if (
                    str(worksheet.cell(r, 0).value).find("Loại") == -1
                    and worksheet.cell(r, 2).value == ""
                    and worksheet.cell(r, 3).value == ""
                    and worksheet.cell(r, 0).value != ""
                ):
                    is_subtitle = 1
                    danh_muc_tai_san = worksheet.cell(r, 1).value
                    continue

                # TH1 loại không có các loại nhỏ bên trong và có luôn số liệu hao mòn và số lượng
                if (
                    str(worksheet.cell(r, 0).value).find("Loại") != -1
                    and worksheet.cell(r, 1).value != ""
                    and worksheet.cell(r, 2).value != ""
                    and worksheet.cell(r, 3).value != ""
                ):
                    response_data.append(
                        {
                            "danh_muc_tai_san": worksheet.cell(r, 1).value,
                            "thoi_gian_su_dung": worksheet.cell(r, 2).value,
                            "ti_le_hao_mon": worksheet.cell(r, 3).value,
                            "loai_tai_san": worksheet.cell(r, 1).value,
                        }
                    )
                    continue

                # TH2 danh mục con có luôn số liệu:
                if (
                    str(worksheet.cell(r, 0).value).find("Loại") == -1
                    and worksheet.cell(r, 2).value != ""
                    and worksheet.cell(r, 3).value != ""
                    and worksheet.cell(r, 0).value != ""
                ):
                    response_data.append(
                        {
                            "danh_muc_tai_san": danh_muc_tai_san_lon,
                            "thoi_gian_su_dung": worksheet.cell(r, 2).value,
                            "ti_le_hao_mon": worksheet.cell(r, 3).value,
                            "loai_tai_san": worksheet.cell(r, 1).value,
                        }
                    )
                    continue

                # TH3 danh mục lớn chưa có thêm đứa con nào
                if is_subtitle == 0:
                    response_data.append(
                        {
                            "danh_muc_tai_san": danh_muc_tai_san_lon,
                            "thoi_gian_su_dung": worksheet.cell(r, 2).value,
                            "ti_le_hao_mon": worksheet.cell(r, 3).value,
                            "loai_tai_san": worksheet.cell(r, 1).value,
                        }
                    )
                    continue

                # TH4 danh mục con có thêm vài thằng con nữa
                if is_subtitle == 1:
                    response_data.append(
                        {
                            "danh_muc_tai_san": danh_muc_tai_san,
                            "thoi_gian_su_dung": worksheet.cell(r, 2).value,
                            "ti_le_hao_mon": worksheet.cell(r, 3).value,
                            "loai_tai_san": worksheet.cell(r, 1).value,
                        }
                    )
            db.commit()
        except:
            return {"messages": "Lỗi không xác định"}, 500
        return {"respone_data": response_data}, 200
