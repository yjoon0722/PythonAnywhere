from openpyxl.cell.read_only import ReadOnlyCell
from openpyxl.cell.read_only import EmptyCell

class OrderData:

    def __init__(self, row = None):        
        self.order_number = ""
        self.product_name = ""
        self.product_quantity = ""
        self.order_name = ""
        self.order_phone = ""
        self.order_mobile_phone = ""
        self.recipient_name = ""
        self.recipient_phone = ""
        self.recipient_mobile_phone = ""
        self.recipient_post_code = ""
        self.recipient_address = ""
        self.invoice_number = ""
        self.delivery_message = ""

        #print(f"cell: {cell} = {type(cell)} = [{cell.row}:{cell.column}] {cell.value}")
        if row is None:
            return

        # Cell
        for cell in row:                
            if type(cell) is EmptyCell:
                continue
            if cell.value is None:
                continue
            
            if cell.column == 1: 
                self.order_number = cell.value                  # 주문번호
            elif cell.column == 2: 
                self.product_name = cell.value                  # 상품명(옵션포함)
            elif cell.column == 3: 
                self.product_quantity = cell.value              # 주문상품수량
            elif cell.column == 4: 
                self.order_name = cell.value                    # 주문자이름
            elif cell.column == 5: 
                self.order_phone = cell.value                   # 주문자전화
            elif cell.column == 6: 
                self.order_mobile_phone = cell.value            # 주문자핸드폰
            elif cell.column == 7: 
                self.recipient_name = cell.value                # 수취인이름
            elif cell.column == 8: 
                self.recipient_phone = cell.value               # 수취인전화
            elif cell.column == 9: 
                self.recipient_mobile_phone = cell.value        # 수취인핸드폰
            elif cell.column == 10:
                self.recipient_post_code = cell.value           # 신)수취인우편번호
            elif cell.column == 11:
                self.recipient_address = cell.value             # 수취인주소
            elif cell.column == 12:
                self.invoice_number = cell.value                # 송장번호
            elif cell.column == 13:
                self.delivery_message = cell.value              # 배송메시지
    
    def __str__(self):
        return str(self.__class__) + ": " + str(self.__dict__)
