function filename = save_image_with_timestamp(~, baseName)
    % Задание 3: Сохранение текущего активного графика в .png с временной меткой
    
    % Запрашиваем точное время у компьютера
    currentTime = datetime('now', 'Format', 'yyyyMMdd_HHmmss');
    currentTimeStr = char(currentTime); 
    
    % Формируем уникальное имя файла
    filename = sprintf('%s_%s.png', baseName, currentTimeStr);
    
    % gcf — это встроенная команда MATLAB, которая берет текущее активное окно.
    % Это решает проблему "неправильных паспортов" окон на 100%!
    print(gcf, filename, '-dpng', '-r300');
    
    fprintf('🖼️  ЗАДАНИЕ 3: График успешно сохранен как: %s\n', filename);
end
