function restoredTable = inverse_scale_data(scaledTable, params)
    % inverse_scale_data - Задание 10: Восстановление изначального масштаба
    % с использованием сохраненных констант из структуры params.
    %
    % Входные параметры:
    %   scaledTable - таблица отмасштабированных ответов (x_new, y_new)
    %   params      - структура с сохраненными коэффициентами из Задания 9

    % Создаем копию таблицы под восстановленные данные
    restoredTable = scaledTable;
    
    % Вытаскиваем обратно векторы x_new и y_new
    x_new = scaledTable.x;
    y_new = scaledTable.y;
    
    % Проверяем, какой метод был применен в Задании 9, и включаем обратную математику
    switch lower(params.method)
        case 'minmax'
            % Восстановление для Формулы 1: x = x_new * (max_x - min_x) + min_x
            restoredTable.x = x_new * (params.max_x - params.min_x) + params.min_x;
            restoredTable.y = y_new * (params.max_y - params.min_y) + params.min_y;
            fprintf('🔄 [Задание 10] Восстановлен исходный масштаб по константам MinMax.\n');
            
        case 'minmax_bounds'
            % Восстановление для Формулы 2: x = ((x_new - a) * (max_x - min_x)) / (b - a) + min_x
            a = params.a;
            b = params.b;
            
            restoredTable.x = ((x_new - a) * (params.max_x - params.min_x)) / (b - a) + params.min_x;
            restoredTable.y = ((y_new - a) * (params.max_y - params.min_y)) / (b - a) + params.min_y;
            fprintf('🔄 [Задание 10] Восстановлен исходный масштаб из границ [%.1f, %.1f].\n', a, b);
            
        case 'z_score'
            % Восстановление для Формулы 3 (Z-score): x = z * sigma + mu
            restoredTable.x = x_new * params.sigma_x + params.mu_x;
            restoredTable.y = y_new * params.sigma_y + params.mu_y;
            fprintf('🔄 [Задание 10] Восстановлен исходный масштаб по параметрам выборочного среднего и СКО.\n');
            
        otherwise
            error('Ошибка: Неизвестный метод для восстановления масштаба!');
    end
end
