from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.cell.read_only import EmptyCell

class InvoiceData:
    
    def __init__(self):
        self.key                = ""
        self.fileName			= ""  # 파일명
        self.accountName        = ""  # 판매처명
        self.warehouse			= ""  # 창고
        self.phoneNumber        = ""  # 모바일
        self.itemName			= ""  # 품목명
        self.trackId			= ""  # 일반송장 번호
        self.carrierURL			= ""  # 택배사 배송 조회 URL
        self.carrierName        = ""  # 택배사 이름
        self.carrierId			= ""  # 택배사 코드
        self.statusId			= ""  # 택배 마지막 상태
        self.statusText			= ""  # 택배 마지막 상태
        self.statusTime			= ""  # 택배 마지막 처리 시간
        self.statusLocation     = ""  # 택배 마지막 처리 장소
    
    def __str__(self):
        return str(self.__class__) + ": " + str(self.__dict__)

    def __eq__(self, other):
        return self.accountName == other.accountName

    def __lt__(self, other):
        return self.accountName < other.accountName