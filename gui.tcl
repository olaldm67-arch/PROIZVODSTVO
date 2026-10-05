#!/usr/bin/wish8.6
# =========================================
# Workshop GUI на Tcl/Tk с поддержкой PNG
# Версия: 0.7.7 
# =========================================

set debug_file "workshop_gui_debug.log"
proc debug {msg} {
    global debug_file
    set fd [open $debug_file a]
    puts $fd "[clock format [clock seconds] -format "%H:%M:%S"] - $msg"
    close $fd
}

debug "=== Запуск gui.tcl ==="

#set FONT_TITLE {"Play" 11}
#set FONT_INFO {"Play" 12}
#set FONT_NORMAL {"Play" 11}
#set FONT_INPUT {"Play" 12}
#set FONT_BOLD {"Play" 11 bold}

set FONT_TITLE {"Sans" 11}
set FONT_INFO {"Sans" 12}
set FONT_NORMAL {"Sans" 11}
set FONT_INPUT {"Sans" 12}
set FONT_BOLD {"Sans" 9 bold}
set FONT_ALERT {"Sans" 20 bold}

set COLOR_BG "#f0f0f0"
set COLOR_FRAME "#e8e8e8"
set COLOR_INPUT "#ffffff"
set COLOR_HEADER "#d0d0d0"
set COLOR_ALERT "#FFBB66"

set DRAWING_WIDTH 1280
set DRAWING_HEIGHT 1080
set DATA_WIDTH [expr {1920 - $DRAWING_WIDTH}]
set TABLE1_HEIGHT 200
set TABLE2_HEIGHT 380
set INFO_HEIGHT 500

debug "Размеры заданы: drawing=${DRAWING_WIDTH}x${DRAWING_HEIGHT}, data=${DATA_WIDTH}"

set barcode_var ""
set png_image ""

### Создание главного окна
catch {wm title . "Workshop"} result; debug "wm title: $result"
catch {wm geometry . "1920x1080+0+0"} result; debug "wm geometry: $result"
catch {wm protocol . WM_DELETE_WINDOW exit_app} result; debug "wm protocol: $result"

### Стили
ttk::style configure Treeview -background white -fieldbackground white -font $FONT_NORMAL
ttk::style configure Treeview.Heading -background $COLOR_HEADER -font $FONT_NORMAL
debug "Стили Treeview настроены"

### Левый фрейм (чертёж)
frame .left -width $DRAWING_WIDTH -height $DRAWING_HEIGHT -bg $COLOR_BG
place .left -x 0 -y 0 -width $DRAWING_WIDTH -height $DRAWING_HEIGHT
debug "Левый фрейм создан"

canvas .left.canvas -width [expr {$DRAWING_WIDTH - 10}] -height [expr {$DRAWING_HEIGHT - 10}] -bg white -highlightthickness 1 -highlightbackground "#cccccc"
place .left.canvas -x 5 -y 5
debug "Холст для чертежа размещён"

### Правый фрейм ( .right )
frame .right -width $DATA_WIDTH -height $DRAWING_HEIGHT -bg $COLOR_FRAME
place .right -x $DRAWING_WIDTH -y 0 -width $DATA_WIDTH -height $DRAWING_HEIGHT
debug "Правый фрейм создан"

## Таблица 1 ( .right.table1 )
frame .right.table1 -width $DATA_WIDTH -height $TABLE1_HEIGHT -bg $COLOR_FRAME
place .right.table1 -x 0 -y 0 -width $DATA_WIDTH -height $TABLE1_HEIGHT

ttk::treeview .right.table1.tree -columns {name width height open} -show headings -height 8 
#-yscrollcommand ".right.table1.scroll set"
.right.table1.tree heading name -text "Наименование" -anchor w
.right.table1.tree heading width -text "Шир." -anchor center
.right.table1.tree heading height -text "Выс." -anchor center
.right.table1.tree heading open -text "Откр." -anchor center
.right.table1.tree column name -width 380 -anchor w
.right.table1.tree column width -width 30 -anchor center
.right.table1.tree column height -width 30 -anchor center
.right.table1.tree column open -width 30 -anchor center

#scrollbar .right.table1.scroll -orient vertical -command ".right.table1.tree yview"
place .right.table1.tree -x 5 -y 5 -width [expr {$DATA_WIDTH - 10}] -height [expr {$TABLE1_HEIGHT - 10}]
#place .right.table1.scroll -x [expr {$DATA_WIDTH - 25}] -y 10 -width 15 -height [expr {$TABLE1_HEIGHT - 20}]
debug "Таблица 1 размещена"

## Таблица 2 ( .right.table2 )
frame .right.table2 -width $DATA_WIDTH -height $TABLE2_HEIGHT -bg $COLOR_FRAME
place .right.table2 -x 0 -y $TABLE1_HEIGHT -width $DATA_WIDTH -height $TABLE2_HEIGHT

ttk::treeview .right.table2.tree -columns {article name length} -show headings -height 28
#-yscrollcommand ".right.table2.scroll set"
.right.table2.tree heading article -text "Артикул" -anchor w
.right.table2.tree heading name -text "Наименование" -anchor w
.right.table2.tree heading length -text "Длина / Кол-во" -anchor center
.right.table2.tree column article -width 60 -anchor w
.right.table2.tree column name -width 360 -anchor w
.right.table2.tree column length -width 100 -anchor center

#scrollbar .right.table2.scroll -orient vertical -command ".right.table2.tree yview"
place .right.table2.tree -x 5 -y 0 -width [expr {$DATA_WIDTH - 10}] -height [expr {$TABLE2_HEIGHT - 10}]
#place .right.table2.scroll -x [expr {$DATA_WIDTH - 25}] -y 10 -width 15 -height [expr {$TABLE2_HEIGHT - 20}]
debug "Таблица 2 размещена"

## Информация и ввод ( .right.info )
frame .right.info -width $DATA_WIDTH -height $INFO_HEIGHT -bg $COLOR_FRAME
place .right.info -x 0 -y [expr {$TABLE1_HEIGHT + $TABLE2_HEIGHT}] -width $DATA_WIDTH -height $INFO_HEIGHT

frame .right.info.text -bg white -relief solid -borderwidth 0
place .right.info.text -x 5 -y 5 -width [expr {$DATA_WIDTH - 10}] -height 490

#text .right.info.text.widget -width 70 -height 13 -font $FONT_INFO -bg white -fg "#000000" -wrap word -state disabled
text .right.info.text.widget -width 62 -height 25 -font $FONT_INFO -fg "#000000" -bg $COLOR_FRAME -wrap word -state disabled -relief solid -borderwidth 0
place .right.info.text.widget -x 3 -y 3

frame .right.info.input -bg $COLOR_FRAME
place .right.info.input -x 352 -y 12 -width [expr {$DATA_WIDTH - 365}] -height 25

label .right.info.input.label -text "Штрих-код:" -font $FONT_BOLD -bg $COLOR_FRAME -fg "#333333"
place .right.info.input.label -x 0 -y 0

entry .right.info.input.entry -textvariable barcode_var -font $FONT_INPUT -bg $COLOR_INPUT -fg "#000000" -relief solid -borderwidth 0 -width 18
place .right.info.input.entry -x 90 -y 0

### Привязка обработки ввода
bind .right.info.input.entry <KeyRelease> {
    set current [%W get]
    if {$current != ""} {
        set last_char [string index $current end]
        debug "KeyRelease: current='$current', last_char='$last_char' (ASCII: [scan $last_char %c])"
        if {$last_char == "\r" || $last_char == "\n"} {
            debug "Обнаружен терминатор, вызываем process_barcode"
            process_barcode
        }
    }
}

### Добавляем обработку Return (Enter), если сканер без Enter
bind .right.info.input.entry <Return> {
    debug "Событие Return (Enter), вызываем process_barcode"
    process_barcode
}

focus .right.info.input.entry
debug "Область ввода размещена"

### ========= ФУНКЦИЯ ЗАГРУЗКИ PNG ==========
proc load_png {path} {
    global png_image
    debug "load_png вызван с путём: $path"
    
    .left.canvas delete all
    
    if {![file exists $path] || [file size $path] == 0} {
        debug "Файл не найден или пуст"
        _draw_stub "Чертеж не найден"
        return 0
    }
    
    ## Загружаем PNG
    if {[catch {
        set png_image [image create photo -format png -file $path]
        .left.canvas create image 0 0 -image $png_image -anchor nw
        debug "PNG загружен успешно"
    } err]} {
        debug "Ошибка загрузки PNG: $err"
        _draw_stub "Ошибка загрузки чертежа"
        return 0
    }
    
    return 1
}

proc _draw_stub {message} {
    .left.canvas create rectangle 10 10 [expr {1080 - 30}] [expr {1080 - 90}] \
        -fill "#f8f8f8" -outline "#cccccc"
    .left.canvas create text [expr {1080 / 2}] [expr {1080 / 2}] \
        -text $message -font {"Play" 16} -fill "#666666"
    debug "Заглушка нарисована: $message"
}

proc update_info {text} {
    ## Заменяем \n на реальные переводы строк
    set text [string map {\\n \n} $text]
    debug "update_info: $text"
    .right.info.text.widget configure -state normal
    .right.info.text.widget delete 1.0 end
    .right.info.text.widget insert end $text
    .right.info.text.widget configure -state disabled
}

## Обновление таблицы 1: принимает список строк, каждая строка - список из 4 полей
proc update_table1 {data} {
    debug "update_table1: получен список, длина = [llength $data]"
    .right.table1.tree delete [.right.table1.tree children {}]
    foreach row $data {
        debug "update_table1: строка = '$row', llength = [llength $row]"
        if {[llength $row] >= 4} {
            .right.table1.tree insert {} end -values $row
        }
    }
}

## Обновление таблицы 2: принимает список строк, каждая строка - список из 3 полей
proc update_table2 {data} {
    debug "update_table2: получен список, длина = [llength $data]"
    .right.table2.tree delete [.right.table2.tree children {}]
    foreach row $data {
        debug "update_table2: строка = '$row', llength = [llength $row]"
        if {[llength $row] >= 3} {
            .right.table2.tree insert {} end -values $row
        }
    }
}

proc show_alert {message} {
    debug "show_alert: $message"
    # Генерируем уникальное имя окна
    set win ".alert_[clock clicks]"
    toplevel $win
    wm geometry $win "550x110+700+400"
    wm overrideredirect $win 1
    wm attributes $win -topmost 1
    frame $win.bg -bg "#e8e8e8" -relief raised -borderwidth 2
    pack $win.bg -fill both -expand 1
    label $win.bg.label -relief sunken -justify center -font {"Sans" 20 bold} -bg "#ff8866" -fg "#000033" -text $message
    pack $win.bg.label -fill both -expand 1
    # Автоматически закрыть окно через 3 секунды
    after 3000 [list destroy $win]
}

proc process_barcode {} {
    global barcode_var
    set value [string trim $barcode_var]
    if {$value ne ""} {
        debug "process_barcode: отправляем '$value'"
        puts $value
        flush stdout
        set barcode_var ""
        focus .right.info.input.entry
    } else {
        debug "process_barcode: barcode_var пуст"
    }
}

proc clear_interface {} {
    debug "clear_interface"
    .left.canvas delete all
    _draw_stub "Ожидание сканирования"
    update_info ""
    update_table1 {}
    update_table2 {}
}

proc set_focus {} {
    debug "set_focus"
    focus .right.info.input.entry
}

proc exit_app {} {
    debug "exit_app"
    puts "EXIT"
    flush stdout
    destroy .
}

proc read_stdin {} {
    if {[eof stdin]} {
        debug "eof stdin, завершаем"
        exit_app
        return
    }
    if {[gets stdin line] < 0} {
        after 50 read_stdin
        return
    }
    debug "получена команда: '$line'"
    set cmd [lindex $line 0]
    set args [lrange $line 1 end]
    switch -- $cmd {
        "load_png"      { load_png [lindex $args 0] }
        "update_info"   { update_info [lindex $args 0] }
        "update_table1" { update_table1 [lindex $args 0] }
        "update_table2" { update_table2 [lindex $args 0] }
        "show_alert"    { show_alert $args }
        "clear"         { clear_interface }
        "set_focus"     { set_focus }
        "exit"          { exit_app }
        default         { debug "неизвестная команда: $cmd" }
    }
    after 50 read_stdin
}

fconfigure stdin -blocking 0 -buffering line
fileevent stdin readable read_stdin
debug "Обработчик stdin установлен"

puts "GUI_READY"
flush stdout
debug "GUI_READY отправлен, вход в главный цикл"
