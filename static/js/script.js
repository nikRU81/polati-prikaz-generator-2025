let punktCounter = 0;
let fioCounter = 0;

// Инициализация при загрузке страницы
window.onload = function() {
    addPunkt();
    addPunkt();
    addFIO();
    
    // Устанавливаем текущую дату
    const today = new Date();
    document.getElementById('day').value = today.getDate();
    document.getElementById('month').value = getMonthName(today.getMonth());
    document.getElementById('year').value = today.getFullYear();
};

function getMonthName(monthIndex) {
    const months = [
        'января', 'февраля', 'марта', 'апреля', 'мая', 'июня',
        'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'
    ];
    return months[monthIndex];
}

function addPunkt() {
    punktCounter++;
    const container = document.getElementById('punktContainer');
    const punktDiv = document.createElement('div');
    punktDiv.className = 'punkt-group';
    punktDiv.id = `punkt-${punktCounter}`;
    
    punktDiv.innerHTML = `
        <button type="button" class="remove-punkt-btn" onclick="removePunkt(${punktCounter})">✖ Удалить</button>
        <h3>Пункт ${punktCounter}</h3>
        <div class="form-group">
            <label for="punktText-${punktCounter}">Текст пункта*</label>
            <textarea id="punktText-${punktCounter}" required placeholder="Назначить Иванова И.И. ответственным за охрану труда с 22 октября 2025 года."></textarea>
        </div>
    `;
    
    container.appendChild(punktDiv);
}

function removePunkt(id) {
    const element = document.getElementById(`punkt-${id}`);
    if (element) {
        element.remove();
    }
    renumberPunkts();
}

function renumberPunkts() {
    const container = document.getElementById('punktContainer');
    const punkts = container.querySelectorAll('.punkt-group');
    punkts.forEach((punkt, index) => {
        const h3 = punkt.querySelector('h3');
        if (h3) {
            h3.textContent = `Пункт ${index + 1}`;
        }
    });
}

function addFIO() {
    fioCounter++;
    const container = document.getElementById('fioContainer');
    const fioDiv = document.createElement('div');
    fioDiv.className = 'fio-item';
    fioDiv.id = `fio-${fioCounter}`;
    
    fioDiv.innerHTML = `
        <input type="text" id="fioText-${fioCounter}" placeholder="Иванов И.И.">
        <button type="button" onclick="removeFIO(${fioCounter})">✖</button>
    `;
    
    container.appendChild(fioDiv);
}

function removeFIO(id) {
    const element = document.getElementById(`fio-${id}`);
    if (element) {
        element.remove();
    }
}

// Функция для показа уведомлений
function showNotification(message, type = 'info') {
    const notification = document.getElementById('notification');
    notification.textContent = message;
    notification.className = `notification ${type}`;
    notification.classList.add('show');
    
    setTimeout(() => {
        notification.classList.remove('show');
    }, 4000);
}

// Обработка формы
document.getElementById('prikazForm').addEventListener('submit', async function(e) {
    e.preventDefault();
    
    // Получаем кнопку и меняем её состояние
    const btn = document.getElementById('generateBtn');
    const btnText = document.getElementById('btnText');
    const btnLoader = document.getElementById('btnLoader');
    
    btn.disabled = true;
    btnText.style.display = 'none';
    btnLoader.style.display = 'inline';
    
    try {
        // Собираем данные
        const day = document.getElementById('day').value;
        const month = document.getElementById('month').value;
        const year = document.getElementById('year').value;
        const orderNumber = document.getElementById('orderNumber').value;
        const orderTitle = document.getElementById('orderTitle').value;
        const preamble = document.getElementById('preamble').value;

        // Собираем пункты
        const punkts = [];
        const punktElements = document.querySelectorAll('.punkt-group');
        punktElements.forEach((element, index) => {
            const textarea = element.querySelector('textarea');
            if (textarea && textarea.value.trim()) {
                punkts.push({
                    number: index + 1,
                    text: textarea.value.trim()
                });
            }
        });

        if (punkts.length === 0) {
            showNotification('Добавьте хотя бы один пункт приказа!', 'error');
            btn.disabled = false;
            btnText.style.display = 'inline';
            btnLoader.style.display = 'none';
            return;
        }

        // Собираем ФИО
        const fios = [];
        const fioElements = document.querySelectorAll('.fio-item input');
        fioElements.forEach(input => {
            if (input.value.trim()) {
                fios.push(input.value.trim());
            }
        });

        // Формируем данные для отправки
        const formData = {
            day,
            month,
            year,
            orderNumber,
            orderTitle,
            preamble,
            punkts,
            fios
        };

        // Отправляем запрос на сервер
        const response = await fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(formData)
        });

        if (!response.ok) {
            const errorData = await response.json();
            throw new Error(errorData.error || 'Ошибка при генерации документа');
        }

        // Получаем файл
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Приказ_ПОЛАТИ_${orderNumber.replace(/\//g, '-')}.docx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        showNotification('✅ Приказ успешно сгенерирован и скачан!', 'success');
    } catch (error) {
        console.error('Ошибка:', error);
        showNotification(`❌ Ошибка: ${error.message}`, 'error');
    } finally {
        // Возвращаем кнопку в нормальное состояние
        btn.disabled = false;
        btnText.style.display = 'inline';
        btnLoader.style.display = 'none';
    }
});
