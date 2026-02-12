<tbody class="divide-y text-sm">
    {% for item in results %}
    <tr class="hover:bg-blue-50 cursor-pointer" 
        data-id="{{ item.row_id }}"
        onclick="selectRow(this, '{{ item.두께 }}', '{{ item.품목 }}', '{{ item.평량 }}', '{{ item.고시가 }}', '{{ item.사이즈 }}', '{{ item.색상 }}')">
        <td class="p-4 text-center">
            <button onclick="event.stopPropagation(); toggleFavorite('{{ item.row_id }}', '{{ item.품목 }}', '{{ item.두께 }}', '{{ item.평량 }}', '{{ item.고시가 }}', '{{ item.사이즈 }}', '{{ item.시트명 }}', '{{ item.색상 }}', this)" 
                    class="fav-star text-gray-300 text-2xl hover:scale-110">★</button>
        </td>
        </tr>
    {% endfor %}
</tbody>

<script>
    function toggleFavorite(id, name, thick, gram, price, size, sheet, color, btn) {
        let favs = JSON.parse(localStorage.getItem('paper_favs') || '[]');
        const idx = favs.findIndex(f => f.id === id);

        if (idx > -1) {
            favs.splice(idx, 1);
        } else {
            favs.push({ id, name, thick, gram, price, size, sheet, color });
        }
        localStorage.setItem('paper_favs', JSON.stringify(favs));
        renderFavorites();
    }

    function renderFavorites() {
        const bar = document.getElementById('favorites_bar');
        const favs = JSON.parse(localStorage.getItem('paper_favs') || '[]');
        
        // 현재 화면의 별표 상태 업데이트
        document.querySelectorAll('tr[data-id]').forEach(tr => {
            const starBtn = tr.querySelector('.fav-star');
            const rowId = tr.getAttribute('data-id');
            if(favs.some(f => f.id === rowId)) starBtn.classList.add('active');
            else starBtn.classList.remove('active');
        });

        bar.innerHTML = favs.map(f => `
            <div class="flex items-center gap-2 px-4 py-2 bg-white border border-yellow-200 rounded-2xl shadow-sm cursor-pointer hover:bg-yellow-50 active:scale-95 transition-all" 
                 onclick="quickSelect('${f.thick}', '${f.name}', '${f.gram}', '${f.price}', '${f.size}', '${f.color}')">
                <span class="text-yellow-400 text-sm">★</span>
                <div class="flex flex-col leading-tight">
                    <span class="text-[11px] font-black text-gray-700">${f.name} ${f.color !== '-' ? '['+f.color+']' : ''} (${f.gram}g)</span>
                    <span class="text-[9px] text-gray-400">${f.size}</span>
                </div>
                <button onclick="event.stopPropagation(); removeFav('${f.id}')" class="text-gray-300 hover:text-red-500 ml-1 text-lg font-bold">×</button>
            </div>
        `).join('');
    }
    // ... 나머지 함수 동일 ...
</script>