function [scaledTable, params] = scale_data(inputTable)
    % scale_data - Задание 9: Интерактивный расчет масштабирования.
    
    scaledTable = inputTable;
    params = struct(); % Сохраняем константы (Требование Задания 9)
    
    % Извлекаем векторы твоего Варианта 7
    x = inputTable.x;
    y = inputTable.y;
    
    % Вычисляем глобальные экстремумы твоего 7-го варианта для диапазона
    min_x = min(x); max_x = max(x);
    min_y = min(y); max_y = max(y);
    
    mu_x = mean(x); sigma_x = std(x);
    mu_y = mean(y); sigma_y = std(y);
    
    % Выводим меню со строгими формулировками по твоему варианту
    fprintf('\n=======================================================\n');
    fprintf('   📈 МЕНЮ МАСШТАБИРОВАНИЯ ДАННЫХ (Задание 9)         \n');
    fprintf('=======================================================\n');
    fprintf(' 1. Формула 1: Диапазон определяется экстремумами моего 7-го варианта\n');
    fprintf(' 2. Формула 2: Нормализация в произвольных границах a, b\n');
    fprintf(' 3. Формула 3: Стандартизация (z-score) через mu и sigma\n');
    fprintf('=======================================================\n');
    
    formula_choice = input('👉 Выберите формулу для расчета (1-3): ');
    
    switch formula_choice
        case 1
            % x_new = (x - min(x)) / (max(x) - min(x))
            scaledTable.x = (x - min_x) / (max_x - min_x);
            scaledTable.y = (y - min_y) / (max_y - min_y);
            
            params.method = 'minmax';
            params.min_x = min_x; params.max_x = max_x;
            params.min_y = min_y; params.max_y = max_y;
            
            fprintf('✅ Расчет по Формуле 1 выполнен на основе min/max твоего варианта.\n');
            
        case 2
            % x_new = a + ((x - min(x))*(b - a)) / (max(x) - min(x))
            fprintf('   Введите произвольные границы масштабирования:\n');
            a_val = input('   👉 Введите произвольную нижнюю границу (a): ');
            b_val = input('   👉 Введите произвольную верхнюю границу (b): ');
            
            scaledTable.x = a_val + ((x - min_x) * (b_val - a_val)) / (max_x - min_x);
            scaledTable.y = a_val + ((y - min_y) * (b_val - a_val)) / (max_y - min_y);
            
            params.method = 'minmax_bounds';
            params.a = a_val; params.b = b_val;
            params.min_x = min_x; params.max_x = max_x;
            params.min_y = min_y; params.max_y = max_y;
            
            fprintf('✅ Расчет по Формуле 2 выполнен в произвольных границах [%.2f, %.2f].\n', a_val, b_val);
            
        case 3
            % z = (x - mu) / sigma
            scaledTable.x = (x - mu_x) / sigma_x;
            scaledTable.y = (y - mu_y) / sigma_y;
            
            params.method = 'z_score';
            params.mu_x = mu_x; params.sigma_x = sigma_x;
            params.mu_y = mu_y; params.sigma_y = sigma_y;
            
            fprintf('✅ Расчет по Формуле 3 выполнен.\n');
            
        otherwise
            warning('⚠️ Неверный выбор! Данные возвращены без масштабирования.');
    end
end