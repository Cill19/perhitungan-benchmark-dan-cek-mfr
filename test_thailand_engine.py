import pandas as pd

from thailand_engine import process_shopee_th


def test_shopee_th_basic_calculation():
    """
    Test basic Shopee Thailand benchmark calculation.

    Expected:
    - Valid orders included
    - Cancelled orders excluded
    - Seller voucher deducted once per order
    """

    data = pd.DataFrame(
        {
            "หมายเลขคำสั่งซื้อ": [
                "ORDER001",
                "ORDER001",
                "ORDER002",
                "ORDER003",
            ],

            "สถานะการสั่งซื้อ": [
                "สำเร็จ",
                "สำเร็จ",
                "ยกเลิกแล้ว",
                "สำเร็จ",
            ],

            "วันที่ทำการสั่งซื้อ": [
                "01/08/2026",
                "01/08/2026",
                "02/08/2026",
                "15/08/2026",
            ],

            "ราคาขายสุทธิ": [
                "98.00",
                "200.00",
                "500.00",
                "300.00",
            ],

            "โค้ดส่วนลดชำระโดยผู้ขาย": [
                "20.00",
                "20.00",
                "50.00",
                "0.00",
            ],
        }
    )


    result = process_shopee_th(data)


    # ORDER001:
    # GMV = 98 + 200 = 298
    # Seller voucher = 20 (NOT 40)
    #
    # ORDER002:
    # Cancelled → excluded
    #
    # ORDER003:
    # GMV = 300
    #
    # Benchmark:
    # (298 - 20) + 300
    # = 578 THB

    total_benchmark = result["benchmark_gmv"].sum()


    assert total_benchmark == 578


def test_thb_decimal_not_converted_to_idr():

    """
    Make sure:
    98.00 THB stays 98
    NOT 9800
    """

    data = pd.DataFrame(
        {
            "หมายเลขคำสั่งซื้อ": [
                "ORDER001"
            ],

            "สถานะการสั่งซื้อ": [
                "สำเร็จ"
            ],

            "วันที่ทำการสั่งซื้อ": [
                "01/08/2026"
            ],

            "ราคาขายสุทธิ": [
                "98.00"
            ],

            "โค้ดส่วนลดชำระโดยผู้ขาย": [
                "0.00"
            ],
        }
    )


    result = process_shopee_th(data)


    assert result["benchmark_gmv"].sum() == 98


def test_cancel_keyword_thailand():

    """
    Ensure Thai cancelled status is excluded.
    """

    data = pd.DataFrame(
        {
            "หมายเลขคำสั่งซื้อ": [
                "ORDER001"
            ],

            "สถานะการสั่งซื้อ": [
                "ยกเลิกแล้ว"
            ],

            "วันที่ทำการสั่งซื้อ": [
                "01/08/2026"
            ],

            "ราคาขายสุทธิ": [
                "1000.00"
            ],

            "โค้ดส่วนลดชำระโดยผู้ขาย": [
                "100.00"
            ],
        }
    )


    result = process_shopee_th(data)


    assert result["benchmark_gmv"].sum() == 0


if __name__ == "__main__":

    test_shopee_th_basic_calculation()
    test_thb_decimal_not_converted_to_idr()
    test_cancel_keyword_thailand()

    print(
        "All Thailand Shopee tests passed."
    )