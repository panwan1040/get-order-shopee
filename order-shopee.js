(async function() {
    // Function to load the xlsx library dynamically
    function loadScript(url) {
        return new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = url;
            script.onload = () => resolve();
            script.onerror = () => reject(new Error(`Script load error for ${url}`));
            document.head.appendChild(script);
        });
    }

    // Load the xlsx library
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.3/xlsx.full.min.js');

    async function fetchWithRetry(url, retries = 3) {
        for (let i = 0; i < retries; i++) {
            try {
                const response = await fetch(url);
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                const json = await response.json();
                console.log(`fetchWithRetry: Success on attempt ${i + 1}`);
                return json;
            } catch (error) {
                console.error(`fetchWithRetry: Attempt ${i + 1} failed:`, error);
                if (i === retries - 1) throw error;
            }
        }
    }

    async function getOrders(offset, limit) {
        const url = `https://shopee.co.th/api/v4/order/get_all_order_and_checkout_list?limit=${limit}&offset=${offset}`;
        console.log(`getOrders: Fetching URL: ${url}`);
        try {
            const result = await fetchWithRetry(url);
            console.log(`getOrders: Fetched data for offset ${offset}:`, result);
            if (!result.data || !result.data.order_data) {
                console.error(`getOrders: Unexpected data structure:`, result);
                return [];
            }
            return result.data.order_data.details_list || [];
        } catch (error) {
            console.error("getOrders: Error fetching orders:", error);
            return [];
        }
    }

    async function getOrderDetails(orderId) {
        const url = `https://shopee.co.th/api/v4/order/get_order_detail?order_id=${orderId}`;
        console.log(`getOrderDetails: Fetching URL: ${url}`);
        try {
            const result = await fetchWithRetry(url);
            console.log(`getOrderDetails: Fetched data for order ${orderId}:`, result);
            const orderInfoRows = result.data.processing_info.info_rows;
            const orderTimeInfo = orderInfoRows.find(row => row.info_label.text === "label_odp_order_time");
            return orderTimeInfo ? orderTimeInfo.info_value.value : null;
        } catch (error) {
            console.error(`getOrderDetails: Error fetching details for order ${orderId}:`, error);
            return null;
        }
    }

    function formatCurrency(number) {
        number = number/1000
        return new Intl.NumberFormat('th-TH', { style: 'currency', currency: 'THB' }).format(number);
    }

    function unixToDate(unixTimestamp) {
        const date = new Date(unixTimestamp * 1000);
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        return `${day}-${month}-${year}`;
    }

    function unixToExcelDate(unixTimestamp) {
        const date = new Date(unixTimestamp * 1000);
        const excelDate = (date - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
        return excelDate;
    }

    async function getAllOrders() {
        const limit = 20;
        let offset = 0;
        const allOrders = [
            ['OrderID', 'Order Time', 'ชื่อสินค้า', 'จำนวน', 'ราคารวม', 'สถานะ', 'ชื่อร้านค้า', 'รายละเอียด', 'ราคาเดิม']
        ];
        let totalAmount = 0;
        let totalCount = 0;

        while (true) {
            console.log(`getAllOrders: Fetching orders with offset ${offset}`);
            const data = await getOrders(offset, limit);
            console.log(`getAllOrders: Orders fetched with offset ${offset}:`, data);
            if (data.length === 0) break;

            const detailsPromises = data.map(async item => {
                const infoCard = item.info_card;
                const listType = item.list_type;
                const orderid = infoCard.order_id;

                console.log(`getAllOrders: Processing order ${orderid}`);

                const statusMap = {
                    3: "เสร็จสิ้น",
                    4: "ยกเลิก",
                    7: "กำลังจัดส่ง",
                    8: "กำลังขนส่ง",
                    9: "รอชำระเงิน",
                    12: "คืนสินค้า"
                };
                const strListType = statusMap[listType] || "ไม่ทราบ";

                const productCount = infoCard.product_count;
                let subTotal = infoCard.subtotal / 1e2; // Adjusted for THB
                if (listType !== 4 && listType !== 12) totalAmount += subTotal;
                else subTotal = 0;
                totalCount += productCount;

                const orderCard = infoCard.order_list_cards[0];
                const { username, shop_name: shopName } = orderCard.shop_info;
                const products = orderCard.product_info.item_groups;
                const productSummary = products.flatMap(group => group.items.map(item => item.model_name)).join(', ');

                const name = products[0].items[0].name;
                const formattedTotal = formatCurrency(subTotal);

                // Fetch the order time
                const orderTime = await getOrderDetails(orderid);
                const formattedOrderTime = orderTime ? unixToDate(orderTime) : '';

                console.log(`getAllOrders: Order ${orderid} processed:`, {
                    orderid, formattedOrderTime, name, productCount, formattedTotal, strListType, shopName, productSummary, subTotal
                });

                return [
                    orderid, formattedOrderTime, name, productCount, formattedTotal, strListType, `${username} - ${shopName}`, productSummary, formatCurrency(subTotal)
                ];
            });

            const orderDetails = await Promise.all(detailsPromises);
            allOrders.push(...orderDetails);

            console.log(`getAllOrders: Collected orders with offset ${offset}`);
            offset += limit;
        }

        allOrders.push([
            'รวมทั้งหมด: ', '', '', totalCount, formatCurrency(totalAmount), '', '', '', ''
        ]);

        // Create a workbook and add the data
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(allOrders);
        XLSX.utils.book_append_sheet(wb, ws, 'Orders');

        // Generate the Excel file and download it
        XLSX.writeFile(wb, 'shopee_orders.xlsx');
    }

    getAllOrders();
})();
