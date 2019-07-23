from run import db


class AssetModel(db.Model):
    __tablename__ = "asset"

    id = db.Column(db.Integer, primary_key=True)
    stt = db.Column(db.Integer)
    danh_muc_tai_san = db.Column(db.String(200))
    thoi_gian_su_dung = db.Column(db.Float)
    ti_le_hao_mon = db.Column(db.Float)

    def __init__(self, stt, danh_muc_tai_san, thoi_gian_su_dung, ti_le_hao_mon):
        self.stt = stt,
        self.danh_muc_tai_san = danh_muc_tai_san,
        self.thoi_gian_su_dung = thoi_gian_su_dung,
        self.ti_le_hao_mon = ti_le_hao_mon,

    def json(self, stt, danh_muc_tai_san, thoi_gian_su_dung, ti_le_hao_mon):
        return {"stt": self.stt,
                "danh_muc_tai_san": self.danh_muc_tai_san,
                "thoi_gian_su_dung": self.thoi_gian_su_dung,
                "ti_le_hao_mon": self.ti_le_hao_mon}
