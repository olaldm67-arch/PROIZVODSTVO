#!/bin/bash
# ============================================
# Workshop - автоматизация производства ROSHAN
# Версия: 0.7.7 Закрытие старых смен
# ============================================

### ====== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ И НАСТРОЙКИ =======

SQL_CONNECT='mysql -h 192.168.0.10 -u workshop --password=w0rK5h0p -D workshop --ssl=0 -Nse'
WSHP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_FILE="$WSHP_DIR/workshop.log"
EMF_PATH="/mnt/smb/ПВХ/ЗАКАЗЫ"
VIEW_PNG="$WSHP_DIR/VIEW.png"

GUI_SCRIPT="$WSHP_DIR/gui.tcl"

# Глобальные переменные
declare -g myid time_out wsgroup wshost
declare -g time_start_s=0
declare -g dt="$(date +'%F %T')"

declare -g zakaz="none"
declare -g konname="none"
declare -g konid="none"
declare -g my_info="none"
declare -g kcolor="none"
declare -g work_type="none"
declare -g obj_id="none"
declare -g zak_id="none"

declare -g sql_table1=""
declare -g sql_table2=""
declare -g detal_cur=""
declare -g shtrih_clean=""
declare -g EMF=""

declare -g wkr_max=0          # максимальное количество работников на рабочем месте (RAB_MAX)
declare -g wkr_cur=0           # текущее количество (RAB_CUR)

# Переменные для coproc
declare -g GUI_PID
declare -g GUI_IN_FD
declare -g GUI_OUT_FD

### ============= ФУНКЦИИ УТИЛИТ =============

log() {
    local message="$1"
    local level="${2:-INFO}"
    local timestamp=$(date '+%Y-%m-%d %H:%M:%S')
    echo "$timestamp - [$level] - $message" >> "$LOG_FILE"
    
    if [[ "$level" == "ERROR" ]]; then
        echo "ERROR: $message" >&2
    elif [[ "$level" == "WARNING" ]]; then
        echo "WARNING: $message" >&2
    fi
}

switch_off() {
    if [ -n "$detal_cur" ]; then
		write_detal_end "$detal_cur"
    fi
	wkr_unreg_all
    send_gui_cmd "show_alert { КОНЕЦ СМЕНЫ }"
	systemctl poweroff
}

error_exit() {
    local module="$1"
    local message="$2"
    log "Ошибка в $module: $message" "ERROR"
    send_gui_cmd "show_alert {ОШИБКА: $message}"
    cleanup
    exit 1
}

check_result() {
    local result=$?
    local operation="$1"
    if [ $result -ne 0 ]; then
        log "Ошибка при выполнении: $operation (код: $result)" "ERROR"
        return $result
    fi
    return 0
}

# Конвертация EMF в PNG
emf_to_png() {
    EMF="${EMF_PATH}/${zakaz}(${konid}).emf"
    log "Конвертация EMF в PNG: $EMF"
    
    if [ ! -f "$EMF" ]; then
        log "Файл EMF не найден: $EMF" "WARNING"
        return 1
    fi
    
    inkscape --export-width=1280 --export-height=1080 --export-type=png --export-filename=- "$EMF" 2>/dev/null | \
    convert png:- -crop 100%x100%+0+50 "$VIEW_PNG" 2>/dev/null || {
        log "Ошибка конвертации EMF в PNG" "ERROR"
        return 1
    }

    log "Конвертация успешно завершена"
    return 0
}

### =========== ФУНКЦИИ РАБОТЫ С GUI ===============

set_focus_to_input() {
    if command -v wmctrl &>/dev/null; then
        local win_id
        win_id=$(wmctrl -l | grep "Workshop" | head -1 | awk '{print $1}')
        if [ -n "$win_id" ]; then
            wmctrl -i -a "$win_id"
            log "Фокус установлен на окно Workshop (wmctrl)"
        else
            log "Окно Workshop не найдено для установки фокуса" "WARNING"
        fi
    else
        log "wmctrl не установлен, фокус не установлен" "WARNING"
    fi
}

init_gui() {
    log "Инициализация Tcl/Tk GUI (coproc)..."

    if command -v wish8.6 &>/dev/null; then
        local wish_cmd="wish8.6"
    elif command -v wish &>/dev/null; then
        local wish_cmd="wish"
    else
        error_exit "GUI" "Tcl/Tk (wish) не установлен"
    fi

    if [ ! -f "$GUI_SCRIPT" ]; then
        error_exit "GUI" "Файл gui.tcl не найден: $GUI_SCRIPT"
    fi

    coproc GUI_PROCESS { $wish_cmd "$GUI_SCRIPT"; }
    GUI_PID=$!

    exec {GUI_IN_FD}>&${GUI_PROCESS[1]}
    exec {GUI_OUT_FD}<&${GUI_PROCESS[0]}

    log "GUI процесс запущен с PID: $GUI_PID, дескрипторы: in=$GUI_IN_FD, out=$GUI_OUT_FD"

    local timeout=10
    local line=""
    log "Ожидание сигнала GUI_READY (таймаут ${timeout}с)..."
    while read -t $timeout -u $GUI_OUT_FD line; do
        if [ "$line" = "GUI_READY" ]; then
            log "GUI успешно инициализирован"
            set_focus_to_input
            return 0
        fi
    done

    log "Таймаут ожидания GUI. Вывод процесса:" "ERROR"
    cat <&$GUI_OUT_FD
    kill $GUI_PID 2>/dev/null
    wait $GUI_PID 2>/dev/null
    error_exit "GUI" "Таймаут ожидания GUI"
}

send_gui_cmd() {
    local cmd="$1"
    if [ -n "$GUI_PID" ] && kill -0 $GUI_PID 2>/dev/null; then
        echo "$cmd" >&$GUI_IN_FD
        log "Отправка в GUI: $cmd"
    else
        log "GUI не запущен, команда '$cmd' не отправлена" "WARNING"
    fi
}

read_barcode_from_gui() {
    local timeout=$1
    local barcode=""
    if read -t $timeout -u $GUI_OUT_FD barcode; then
        echo "$barcode"
        return 0
    fi
    return 1
}

cleanup_gui() {
    if [ -n "$GUI_PID" ]; then
        log "Завершение GUI процесса: $GUI_PID"
        send_gui_cmd "exit"
        sleep 0.5
        kill $GUI_PID 2>/dev/null || true
        wait $GUI_PID 2>/dev/null || true
    fi
    exec {GUI_IN_FD}>&- 2>/dev/null || true
    exec {GUI_OUT_FD}>&- 2>/dev/null || true
    log "Очистка GUI завершена"
}

### ========== ФУНКЦИИ РАБОТЫ С БАЗОЙ =============

check_db_connection() {
    log "Проверка подключения к базе данных..."
    if ! eval "$SQL_CONNECT \"SELECT 1;\"" > /dev/null 2>&1; then
        log "Нет подключения к базе данных" "ERROR"
        return 1
    fi
    log "Подключение к базе данных успешно"
    return 0
}

read_start_sql() {
    log "Чтение данных рабочего места из БД..."

    cur_ip=$(hostname -I | awk '{print $1}')
    myid=$(echo "$cur_ip" | awk -F'.' '{print $NF}')
    #myid=201
	
    log "Определен myid: $myid (из IP: $cur_ip)"
    
    time_out=$(eval "$SQL_CONNECT \"SELECT MIN_T FROM WORKPLACE WHERE IP=$myid;\"")
    check_result "SELECT MIN_T" || return 1
    
    wsgroup=$(eval "$SQL_CONNECT \"SELECT GR FROM WORKPLACE WHERE IP=$myid;\"")
    check_result "SELECT GR" || return 1
    
    wshost=$(eval "$SQL_CONNECT \"SELECT NAME FROM WORKPLACE WHERE IP=$myid;\"")
    check_result "SELECT NAME" || return 1
    
    ## Новые поля для работников
    wkr_max=$(eval "$SQL_CONNECT \"SELECT RAB_MAX FROM WORKPLACE WHERE IP=$myid;\"")
    check_result "SELECT RAB_MAX" || return 1
    
    ## Сбрасываем текущее количество работников при старте (на случай перезапуска)
    local reset_sql="UPDATE WORKPLACE SET RAB_CUR=0 WHERE IP=$myid;"
    eval "$SQL_CONNECT \"$reset_sql\""
    check_result "Сброс RAB_CUR" || return 1
    wkr_cur=0

    ## Закрываем прошлые(вчерашние, пятничные) смены
    close_old_session="UPDATE WORKER_TIME SET TIME_END_W = CONCAT(DATE(TIME_START_W),' 18:00:00') WHERE TIME_END_W IS NULL AND DATE(TIME_START_W) < CURDATE();"
    eval "$SQL_CONNECT \"$close_old_session\""
    check_result "Старые смены закрыты" || return 1

    #fio_list=$(eval "$SQL_CONNECT \"SELECT w.FIO FROM WORKER_TIME wt INNER JOIN WORKER w ON wt.RAB=w.RAB WHERE wt.IP=$myid AND wt.TIME_END_W IS NULL;\"")
	#fio_list=$(eval "$SQL_CONNECT \"SELECT GROUP_CONCAT(w.FIO ORDER BY w.FIO SEPARATOR ', ') AS result FROM WORKER_TIME wt INNER JOIN WORKER w ON wt.RAB=w.RAB WHERE wt.IP=$myid AND wt.TIME_END_W IS NULL;\"")
    #check_result "SELECT w.FIO" || return 1
    
	log "Данные рабочего места загружены: $wshost (IP: $myid, группа: $wsgroup, таймаут: $time_out, макс. работников: $wkr_max)"
    return 0
}

read_zakaz_sql() {
    log "Чтение данных заказа для детали: $shtrih_clean"
    
    if [ -z "$shtrih_clean" ]; then
        log "Пустой штрих-код для чтения заказа" "ERROR"
        return 1
    fi
    
    zakaz=$(eval "$SQL_CONNECT \"SELECT zaknum FROM CCALC WHERE CCALC.DETAL='$shtrih_clean';\"")
    check_result "SELECT zaknum" || return 1
    
    konname=$(eval "$SQL_CONNECT \"SELECT konname FROM CCALC WHERE CCALC.DETAL='$shtrih_clean';\"")
    check_result "SELECT konname" || return 1
    
    konid=$(eval "$SQL_CONNECT \"SELECT konid FROM CCALC WHERE CCALC.DETAL='$shtrih_clean';\"")
    check_result "SELECT konid" || return 1
    
    my_info=$(eval "$SQL_CONNECT \"SELECT PRIMPROIZV FROM CCALC WHERE CCALC.DETAL='$shtrih_clean';\"")
    check_result "SELECT PRIMPROIZV" || return 1
    
    kcolor=$(eval "$SQL_CONNECT \"SELECT matcolor FROM CCALC WHERE CCALC.DETAL='$shtrih_clean';\"")
    check_result "SELECT matcolor" || return 1
    
    work_type=$(eval "$SQL_CONNECT \"SELECT GR FROM WORKPLACE WHERE WORKPLACE.IP='$myid';\"")
    check_result "SELECT GR" || return 1
    
    obj_id=$(eval "$SQL_CONNECT \"SELECT OBJID FROM CCALC WHERE CCALC.DETAL='$shtrih_clean';\"")
    check_result "SELECT OBJID" || return 1
    
    zak_id=$(eval "$SQL_CONNECT \"SELECT ZAKID FROM CCALC WHERE CCALC.DETAL='$shtrih_clean';\"")
    check_result "SELECT ZAKID" || return 1
    
    sql_table1=$(eval "$SQL_CONNECT \"SELECT LEFT(REPLACE(CASE WHEN LEFT(NAME,1)='*' THEN MID(NAME,3) ELSE NAME END,' ','_'),50) AS NAME, ROUND(L,0) AS L, ROUND(L,0) AS H, NAPRNAME FROM CCALC_FRN WHERE ZAKID='$zak_id' AND KONID='$konid';\"")
    check_result "SELECT CCALC_FRN" || return 1

    sql_table2=$(eval "$SQL_CONNECT \"SELECT LEFT(REPLACE(CASE WHEN LEFT(MATART,1)='*' THEN MID(MATART,3) ELSE MATART END,' ','_'),50) AS MATART, LEFT(REPLACE(CASE WHEN LEFT(MATNAME,1)='*' THEN MID(MATNAME,3) ELSE MATNAME END,' ','_'),50) AS MATNAME, ROUND(L,0) AS L FROM CCALC_TBL WHERE ZAKID='$zak_id' AND KONID='$konid' AND OBJID='$obj_id' AND GR='$work_type';\"")
    check_result "SELECT CCALC_TBL" || return 1

	#fio_list=$(eval "$SQL_CONNECT \"SELECT w.FIO FROM WORKER_TIME wt INNER JOIN WORKER w ON wt.RAB=w.RAB WHERE wt.IP=$myid AND wt.TIME_END_W IS NULL;\"")
    #fio_list=$(eval "$SQL_CONNECT \"SELECT GROUP_CONCAT(w.FIO ORDER BY w.FIO SEPARATOR ', ') AS result FROM WORKER_TIME wt INNER JOIN WORKER w ON wt.RAB=w.RAB WHERE wt.IP=$myid AND wt.TIME_END_W IS NULL;\"")
    #check_result "SELECT w.FIO" || return 1
    
    log "Данные заказа загружены: Заказ: $zakaz, Контрагент: $konname, ID: $konid"
    return 0
}

write_detal_time() {
    local shtrih_clean="$1"
    log "Запись времени начала обработки детали: $shtrih_clean"
    sql_insert="INSERT INTO DETAL_TIME (IP, DETAL, GR, TIME_START_D) VALUES ($myid, '$shtrih_clean', $wsgroup, '$dt');"
    eval "$SQL_CONNECT \"$sql_insert\""
    check_result "INSERT DETAL_TIME"
}

chek_detal_end() {
    log "Проверка завершения обработки детали: $shtrih_clean"
    
    if [ -z "$shtrih_clean" ]; then
        log "Пустой штрих-код для проверки" "ERROR"
        return 1
    fi
    
    local query="SELECT EXISTS (SELECT 1 FROM DETAL_TIME WHERE DETAL='$shtrih_clean' AND GR=$wsgroup);"
    chek_end=$(eval "$SQL_CONNECT \"$query\"")
    check_result "SELECT EXISTS DETAL_TIME"
    
    log "Результат проверки детали $shtrih_clean: $chek_end"
    return $chek_end
}

write_detal_end() {
    local shtrih_end="$1"
    log "Запись времени окончания обработки детали: $shtrih_end"
    sql_update="UPDATE DETAL_TIME SET TIME_END_D='$dt' WHERE IP=$myid AND DETAL='$shtrih_end';"
    eval "$SQL_CONNECT \"$sql_update\""
    check_result "UPDATE DETAL_TIME"
}

sql_delete() {
    local shtrih_clean="$1"
    log "Удаление записи о детали: $shtrih_clean"
    sql_delete="DELETE FROM DETAL_TIME WHERE IP=$myid AND DETAL='$shtrih_clean';"
    eval "$SQL_CONNECT \"$sql_delete\""
    check_result "DELETE DETAL_TIME"
}

## ------- Добавление работников -------------------
wkr_check() {
    local wkr_id="$1"
    local sql="SELECT IP FROM WORKER_TIME WHERE RAB='$wkr_id' AND TIME_END_W IS NULL AND DATE(TIME_START_W) = CURDATE();"
    local result=$(eval "$SQL_CONNECT \"$sql\"")
    check_result "Проверка регистрации работника $wkr_id" || return 1
    echo "$result"  # возвращаем IP (может быть пустым)
}

## Регистрация работника на текущем рабочем месте
wkr_adreg() {
    local wkr_id="$1"
    local dt=$(date +'%F %T')
    
    # Вставляем запись о начале смены
    local sql_insert="INSERT INTO WORKER_TIME (RAB, IP, TIME_START_W) VALUES ('$wkr_id', '$myid', '$dt');"
    eval "$SQL_CONNECT \"$sql_insert\""
    check_result "Регистрация работника $wkr_id" || return 1
    
    # Увеличиваем счётчик на рабочем месте
    local sql_update="UPDATE WORKPLACE SET RAB_CUR = RAB_CUR + 1 WHERE IP = $myid;"
    eval "$SQL_CONNECT \"$sql_update\""
    check_result "Увеличение RAB_CUR" || return 1
    
    # Обновляем локальную переменную (для логов пригодится)
    wkr_cur=$((wkr_cur + 1))
    log "Работник $wkr_id зарегистрирован на месте $myid (теперь $wkr_cur из $wkr_max)"
}

## Снятие регистрации работника (завершение смены)
wkr_unreg() {
    local wkr_id="$1"
    local ip="$2"   # IP рабочего места, где он сейчас
    local dt=$(date +'%F %T')
    
    # Закрываем открытую запись
    local sql_update="UPDATE WORKER_TIME SET TIME_END_W='$dt' WHERE RAB='$wkr_id' AND TIME_END_W IS NULL AND TIME_START_W > CURDATE();"
    eval "$SQL_CONNECT \"$sql_update\""
    check_result "Снятие регистрации работника $wkr_id" || return 1
    
    # Уменьшаем счётчик на том рабочем месте, откуда он уходит
    local sql_decr="UPDATE WORKPLACE SET RAB_CUR = RAB_CUR - 1 WHERE IP = $ip;"
    eval "$SQL_CONNECT \"$sql_decr\""
    check_result "Уменьшение RAB_CUR" || return 1
    
    # Если это текущее рабочее место, обновим локальную переменную
    if [ "$ip" = "$myid" ]; then
        wkr_cur=$((wkr_cur - 1))
    fi
    log "Работник $wkr_id снят с места $ip"
}

wkr_unreg_all() {
	local sql_update="UPDATE WORKER_TIME SET TIME_END_W=NOW() WHERE IP=$myid AND TIME_END_W IS NULL;"
    eval "$SQL_CONNECT \"$sql_update\""
    check_result "Закрытие всех смен" || return 1
    
}

### ============ БИЗНЕС-ЛОГИКА =============

reset_variables() {
    log "Сброс переменных состояния"
    time_start_s=0
    zakaz="none"
    konname="none"
    konid="none"
    my_info="none"
    kcolor="none"
    work_type="none"
    obj_id="none"
    zak_id="none"
    sql_table1=""
    sql_table2=""
    detal_cur=""
    shtrih_clean=""
    EMF=""
}

check_input() {
    #local shtrih_code="$1"
    #shtrih_clean=$shtrih_code
	shtrih_clean=$1
    
    if [ -z "$shtrih_clean" ]; then
        log "Пустой ввод" "WARNING"
        return
    fi
    
    if [ ${#shtrih_clean} -lt 3 ]; then
        log "Слишком короткий штрих-код: $shtrih_clean" "WARNING"
        send_gui_cmd "show_alert { НЕКОРРЕКТНЫЙ ШТРИХКОД }"
        return
    fi
    
    case "$shtrih_clean" in
        12345*)
			switch_off
            ;;

        22*)
			### -----------------------------------------------
            log "Обработка штрих-кода работника: $shtrih_clean"
            local wkr_id="$shtrih_clean"
            
            # Проверяем, где сейчас зарегистрирован этот работник
            local wkr_ip=$(wkr_check "$wkr_id")
            local check_res=$?
            if [ $check_res -ne 0 ]; then
                log "Ошибка при проверке регистрации работника" "ERROR"
                send_gui_cmd "show_alert { ОШИБКА БД }"
                return
            fi
            
            # Получаем текущее количество работников на нашем месте
            local cur_count=$(eval "$SQL_CONNECT \"SELECT RAB_CUR FROM WORKPLACE WHERE IP=$myid;\"")
            check_result "SELECT RAB_CUR" || return 1
            
            if [ -n "$wkr_ip" ]; then
                # Работник уже где-то зарегистрирован
                if [ "$wkr_ip" = "$myid" ]; then
                    # Он на этом же месте – значит, хочет выйти
                    log "Работник $wkr_id выходит с места $myid"
                    wkr_unreg "$wkr_id" "$myid"
                    send_gui_cmd "show_alert { ВЫХОД РАБОТНИКА }"
					update_gui_interface
                else
                    # Он на другом месте – сначала снимаем его оттуда, потом регистрируем здесь (если есть места)
                    log "Работник $wkr_id сейчас на месте $wkr_ip, перемещаем"
                    
                    # Проверяем, есть ли места у нас
                    if [ $cur_count -lt $wkr_max ]; then
                        # Сначала снимаем с того места
                        wkr_unreg "$wkr_id" "$wkr_ip"
                        # Затем регистрируем здесь
                        wkr_adreg "$wkr_id"
                        send_gui_cmd "show_alert { РАБОТНИК ПЕРЕМЕЩЁН }"
						update_gui_interface
                    else
                        send_gui_cmd "show_alert { !!! МЕСТ НЕТ !!! }"
                    fi
                fi
            else
                # Работник нигде не зарегистрирован сегодня
                if [ $cur_count -lt $wkr_max ]; then
                    log "Регистрация нового работника $wkr_id"
                    wkr_adreg "$wkr_id"
                    send_gui_cmd "show_alert { РАБОТНИК ЗАРЕГИСТРИРОВАН }"
                    update_gui_interface
                else
                    send_gui_cmd "show_alert { !!! МЕСТ НЕТ !!! }"
                fi
            fi
            ;;

        016*)
            log "Обработка штрих-кода изделия: $shtrih_clean"
            local current_time=$(date +%s)
            
            # if [ $time_start_s -eq 0 ]; then
            log "Начало обработки новой детали"
            chek_detal_end
            local check_result=$?
                
            if [ $check_result -eq 0 ]; then
                time_start_s=$current_time
                    if read_zakaz_sql; then
                        write_detal_time "$shtrih_clean"
                        detal_cur="$shtrih_clean"
                        log "Успешная регистрация изделия: $detal_cur"
                        
                        emf_to_png
                        update_gui_interface
                    else
                        log "Ошибка чтения данных заказа" "ERROR"
                        send_gui_cmd "show_alert { ОШИБКА ЧТЕНИЯ ЗАКАЗА }"
                        time_start_s=0
                fi
            else
                    log "Попытка повторной обработки изделия: $shtrih_clean" "WARNING"
                    send_gui_cmd "show_alert { ИЗДЕЛИЕ УЖЕ ОБРАБОТАНО }"
            fi
            elif [ $((current_time - time_start_s)) -gt $time_out ]; then
                log "Завершение обработки по таймауту для детали: $detal_cur"
                if [ -n "$detal_cur" ]; then
                    write_detal_end "$detal_cur"
                fi
                reset_variables
                
                time_start_s=$(date +%s)
                if read_zakaz_sql; then
                    write_detal_time "$shtrih_clean"
                    detal_cur="$shtrih_clean"
                    log "Успешная регистрация изделия после таймаута: $detal_cur"
                    emf_to_png
                    update_gui_interface
                else
                    log "Ошибка чтения данных заказа" "ERROR"
                    send_gui_cmd "show_alert { ОШИБКА ЧТЕНИЯ ЗАКАЗА }"
                    reset_variables
                fi
            else
                log "Попытка двойной регистрации: $shtrih_clean" "WARNING"
                send_gui_cmd "show_alert { ДВОЙНАЯ РЕГИСТРАЦИЯ ИЗДЕЛИЯ }"
            fi
            ;;

        *)
            log "Неизвестный формат штрих-кода: $shtrih_clean" "WARNING"
            send_gui_cmd "show_alert { НЕИЗВЕСТНЫЙ ФОРМАТ ШТРИХКОДА }"
            return
            ;;
    esac
    
    set_focus_to_input
}

update_gui_interface() {
    fio_list=$(eval "$SQL_CONNECT \"SELECT GROUP_CONCAT(REPEAT('- ',10), w.FIO ORDER BY w.FIO SEPARATOR '\n ') AS result FROM WORKER_TIME wt INNER JOIN WORKER w ON wt.RAB=w.RAB WHERE wt.IP=$myid AND wt.TIME_END_W IS NULL;\"")
    check_result "SELECT w.FIO" || return 1
    local list_zakaz=" Участок: $wshost\n\n Заказ: $zakaz\n\n Конструкция: $konname\n\n Цвет: $kcolor\n\n Примечание: $my_info\n\n $fio_list"
    send_gui_cmd "update_info {$list_zakaz}"
    
    # Таблица 1: формируем список списков
    if [ -n "$sql_table1" ]; then
        local table1_data=""
        while IFS=$'\t' read -r name width height open; do
            # Каждая строка оборачивается в фигурные скобки, затем они объединяются через пробел
            table1_data="$table1_data {$name $width $height $open}"
        done <<< "$sql_table1"
        # Отправляем как один аргумент, содержащий весь список (ещё одни фигурные скобки снаружи)
        send_gui_cmd "update_table1 {$table1_data}"
    else
        send_gui_cmd "update_table1 {}"
    fi
    
    # Таблица 2
    if [ -n "$sql_table2" ]; then
        local table2_data=""
        while IFS=$'\t' read -r article name length; do
            table2_data="$table2_data {$article $name $length}"
        done <<< "$sql_table2"
        send_gui_cmd "update_table2 {$table2_data}"
    else
        send_gui_cmd "update_table2 {}"
    fi
    
    if [ -f "$VIEW_PNG" ]; then
        send_gui_cmd "load_png {$VIEW_PNG}"
    fi
    
    send_gui_cmd "set_focus"
    set_focus_to_input
}

work_in() {
    log "Запуск основного рабочего цикла с GUI"
    
    while true; do
        if barcode=$(read_barcode_from_gui 1); then
            if [ "$barcode" = "EXIT" ]; then
                log "Получена команда EXIT из GUI"
                break
            fi
            check_input "$barcode"
        fi
        sleep 0.05
    done
}

cleanup() {
    log "Завершение работы, очистка..."
    
    if [ -n "$detal_cur" ] && [ $time_start_s -ne 0 ]; then
        log "Завершение обработки детали при выходе: $detal_cur"
        write_detal_end "$detal_cur"
    fi
    
    cleanup_gui
    log "Работа завершена"
}

### ================ ОСНОВНОЙ КОД =====================

main() {
    log "=== Запуск приложения Workshop с Tcl/Tk GUI (coproc) ==="
    
    if ! check_db_connection; then
        error_exit "DB" "Нет подключения к базе данных"
    fi
    
    if ! read_start_sql; then
        error_exit "DB" "Ошибка загрузки данных рабочего места"
    fi
    
    init_gui
    
    trap 'log "Получен сигнал завершения"; cleanup; exit 0' SIGINT SIGTERM
    trap 'log "Критическая ошибка: $BASH_COMMAND"; cleanup; exit 1' ERR
    
    reset_variables
    update_gui_interface
    work_in
    
    cleanup
}

log "========================================="
log "Запуск Workshop с Tcl/Tk GUI"
log "Версия: 0.7.7"
log "Дата: $(date)"
log "Директория: $WSHP_DIR"
log "========================================="

main "$@"
