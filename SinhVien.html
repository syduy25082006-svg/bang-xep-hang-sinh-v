<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ thống Quản lý & Trích xuất Tệp</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body class="bg-gray-50 min-h-screen p-5">

    <div class="max-w-6xl mx-auto">
        <div class="bg-white p-6 rounded-2xl shadow-sm border border-gray-200 mb-6">
            <h1 class="text-xl font-bold text-gray-800 mb-4 flex items-center">
                <span class="bg-blue-600 w-2 h-6 rounded mr-2"></span> 
                Trích xuất dữ liệu từ Tệp (Excel/CSV)
            </h1>
            
            <div class="flex flex-wrap gap-4 items-end">
                <div class="flex-1 min-w-[200px]">
                    <label class="block text-sm font-medium text-gray-600 mb-1">Chọn tệp từ máy tính:</label>
                    <input type="file" id="fileUpload" accept=".xlsx, .xls, .csv" 
                        class="w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100 border rounded-lg p-1"/>
                </div>
                <button onclick="exportData()" class="bg-gray-800 text-white px-5 py-2.5 rounded-lg font-medium hover:bg-black transition">
                    📥 Xuất Excel hiện tại
                </button>
            </div>
            <p class="text-xs text-gray-400 mt-3">* Hệ thống sẽ tự đọc các cột: Họ tên (Chữ), MSV (Số/Chữ), Điểm (Số).</p>
        </div>

        <div id="editArea" class="hidden bg-orange-50 p-4 rounded-xl border border-orange-200 mb-6 flex gap-3 items-center">
            <span class="text-orange-700 font-bold text-sm">ĐANG SỬA:</span>
            <input type="text" id="editName" placeholder="Tên" class="p-2 rounded border border-orange-300 text-sm">
            <input type="text" id="editMSV" placeholder="MSV" class="p-2 rounded border border-orange-300 text-sm">
            <input type="number" id="editScore" placeholder="Điểm" class="p-2 rounded border border-orange-300 text-sm w-20">
            <button onclick="confirmEdit()" class="bg-orange-600 text-white px-4 py-2 rounded-lg text-sm font-bold">Lưu thay đổi</button>
            <button onclick="cancelEdit()" class="text-gray-500 text-sm">Hủy</button>
        </div>

        <div class="bg-white rounded-2xl shadow-sm border border-gray-200 overflow-hidden">
            <table class="w-full text-left border-collapse">
                <thead>
                    <tr class="bg-gray-50 border-b border-gray-200">
                        <th class="p-4 font-semibold text-gray-600 w-16 text-center">STT</th>
                        <th class="p-4 font-semibold text-gray-600">Họ và Tên (Chữ)</th>
                        <th class="p-4 font-semibold text-gray-600">Mã Sinh Viên (Số)</th>
                        <th class="p-4 font-semibold text-gray-600 text-center">Điểm số</th>
                        <th class="p-4 font-semibold text-gray-600 text-center">Thao tác</th>
                    </tr>
                </thead>
                <tbody id="tableBody" class="divide-y divide-gray-100 text-gray-700">
                    </tbody>
            </table>
            <div id="emptyState" class="p-10 text-center text-gray-400">
                Chưa có dữ liệu. Vui lòng chọn tệp để hiển thị chữ và số.
            </div>
        </div>
    </div>

    <script>
        let dataList = JSON.parse(localStorage.getItem('myFileData')) || [];
        let currentEditId = null;

        document.getElementById('fileUpload').addEventListener('change', function(e) {
            const file = e.target.files[0];
            const reader = new FileReader();

            reader.onload = function(evt) {
                const workbook = XLSX.read(new Uint8Array(evt.target.result), {type: 'array'});
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const json = XLSX.utils.sheet_to_json(sheet, {header: 1}); // Đọc dạng mảng để dễ bắt lỗi

                // Bỏ qua dòng tiêu đề, duyệt từ dòng thứ 2
                const newItems = json.slice(1).map(row => {
                    if(row.length < 2) return null;
                    return {
                        id: Date.now() + Math.random(),
                        name: row[0] || "N/A",  // Cột 1: Thường là tên (Chữ)
                        msv: row[1] || "000",   // Cột 2: Thường là MSV (Số)
                        score: parseFloat(row[2]) || 0 // Cột 3: Điểm (Số)
                    };
                }).filter(i => i !== null);

                dataList = [...dataList, ...newItems];
                render();
            };
            reader.readAsArrayBuffer(file);
        });

        function render() {
            // Sắp xếp tự động theo điểm cao nhất
            dataList.sort((a, b) => b.score - a.score);
            localStorage.setItem('myFileData', JSON.stringify(dataList));

            const container = document.getElementById('tableBody');
            const empty = document.getElementById('emptyState');
            container.innerHTML = '';

            if(dataList.length > 0) empty.classList.add('hidden');
            else empty.classList.remove('hidden');

            dataList.forEach((item, index) => {
                const row = `
                    <tr class="hover:bg-blue-50/50 transition">
                        <td class="p-4 text-center text-gray-400 font-mono">${index + 1}</td>
                        <td class="p-4 font-medium text-gray-900">${item.name}</td>
                        <td class="p-4 text-blue-600 font-mono">${item.msv}</td>
                        <td class="p-4 text-center font-bold ${item.score >= 5 ? 'text-green-600' : 'text-red-500'}">${item.score}</td>
                        <td class="p-4 text-center">
                            <button onclick="startEdit(${item.id})" class="text-blue-600 hover:underline mr-3 text-sm">Sửa</button>
                            <button onclick="removeItem(${item.id})" class="text-red-400 hover:text-red-600 text-sm">Xóa</button>
                        </td>
                    </tr>
                `;
                container.innerHTML += row;
            });
        }

        // Chỉnh sửa trực tiếp
        function startEdit(id) {
            currentEditId = id;
            const item = dataList.find(i => i.id === id);
            document.getElementById('editName').value = item.name;
            document.getElementById('editMSV').value = item.msv;
            document.getElementById('editScore').value = item.score;
            document.getElementById('editArea').classList.remove('hidden');
        }

        function confirmEdit() {
            const item = dataList.find(i => i.id === currentEditId);
            item.name = document.getElementById('editName').value;
            item.msv = document.getElementById('editMSV').value;
            item.score = parseFloat(document.getElementById('editScore').value) || 0;
            
            cancelEdit();
            render();
        }

        function cancelEdit() {
            document.getElementById('editArea').classList.add('hidden');
            currentEditId = null;
        }

        function removeItem(id) {
            if(confirm("Xóa dòng này?")) {
                dataList = dataList.filter(i => i.id !== id);
                render();
            }
        }

        function exportData() {
            const ws = XLSX.utils.json_to_sheet(dataList.map(i => ({ "Họ Tên": i.name, "MSV": i.msv, "Điểm": i.score })));
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "XepHang");
            XLSX.writeFile(wb, "Du_Lieu_Xep_Hang.xlsx");
        }

        render();
    </script>
</body>
</html>
