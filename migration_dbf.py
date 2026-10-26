import mysql.connector
from dbfread import DBF
from pathlib import Path
import time

DBF_PATH = Path(r"\\sp-roshan-03\Ccalc6.80\Data")

MYSQL_CONFIG = {
    'host': "192.168.0.10",
    'user': "workshop",
    'password': "w0rK5h0p",
    'database': "CCALC",
    'charset': 'utf8mb4',
}

# Только эти таблицы мигрируем (BASE1 первым — источник ZAKNUM)
TABLES = ['BASE1', 'BASE3', 'BASE6', 'BASE7', 'BASE8', 'BASE10', 'BASE14']

BATCH_SIZE = 500


class Migrator:
    def __init__(self, config, path):
        self.config = config
        self.path = path
        self.conn = None

    # ---------- соединение ----------
    def connect(self, retries=3, delay=2):
        for i in range(retries):
            try:
                self.conn = mysql.connector.connect(**self.config)
                print("✓ Подключение к MySQL установлено")
                return True
            except Exception as e:
                print(f"❌ Попытка {i + 1}/{retries}: {e}")
                if i < retries - 1:
                    time.sleep(delay)
        return False

    def ensure(self):
        if not self.conn or not self.conn.is_connected():
            return self.connect()
        return True

    def query(self, sql, params=None):
        if not self.ensure():
            return None
        cur = self.conn.cursor()
        try:
            cur.execute(sql, params or ())
            return cur.fetchall()
        except Exception as e:
            print(f"❌ SQL-запрос: {e}")
            return None
        finally:
            cur.close()

    def execute_many(self, sql, rows):
        if not self.ensure():
            return False
        cur = self.conn.cursor()
        try:
            cur.executemany(sql, rows)
            self.conn.commit()
            return True
        except Exception as e:
            print(f"❌ SQL-вставка: {e}")
            self.conn.rollback()
            return False
        finally:
            cur.close()

    def close(self):
        if self.conn and self.conn.is_connected():
            self.conn.close()

    # ---------- утилиты ----------
    @staticmethod
    def _clean(v):
        if v is None:
            return None
        if isinstance(v, (bytes, bytearray)):
            return None
        if isinstance(v, str):
            return v.strip()[:65535]
        return v

    def find_dbf(self, table_name):
        """Ищет файл таблицы в папке независимо от регистра."""
        for ext in ('*.DBF', '*.dbf'):
            p = self.path / f"{table_name}{ext[1:]}"
            if p.exists():
                return p
        return None

    # ---------- работа с ZAKNUM ----------
    def get_dbf_zaknums(self, dbf_file):
        dbf = DBF(dbf_file, encoding='cp866', ignore_missing_memofile=True)
        zak_field = next((f.name for f in dbf.fields if f.name.upper() == 'ZAKNUM'), None)
        if not zak_field:
            return set()
        return {str(rec.get(zak_field)).strip() for rec in dbf
                if rec.get(zak_field) not in (None, '')}

    def get_mysql_zaknums(self):
        rows = self.query("SELECT ZAKNUM FROM BASE1 WHERE ZAKNUM IS NOT NULL")
        return {str(r[0]).strip() for r in rows} if rows else set()

    # ---------- миграция ----------
    def migrate_table(self, table_name, dbf_file, new_zaknums):
        dbf = DBF(dbf_file, encoding='cp866', ignore_missing_memofile=True)
        fields = [f.name for f in dbf.fields]
        zak_field = next((f for f in fields if f.upper() == 'ZAKNUM'), None)
        if not zak_field:
            print(f"  ⚠️  {table_name}: нет поля ZAKNUM — пропуск")
            return

        cols = ", ".join(f"`{f}`" for f in fields)
        ph = ", ".join(["%s"] * len(fields))
        sql = f"INSERT INTO `{table_name}` ({cols}) VALUES ({ph})"

        batch, count = [], 0
        for rec in dbf:
            zak = rec.get(zak_field)
            if zak is None or str(zak).strip() not in new_zaknums:
                continue
            batch.append(tuple(self._clean(rec.get(f)) for f in fields))
            count += 1
            if len(batch) >= BATCH_SIZE:
                if not self.execute_many(sql, batch):
                    return
                batch = []
        if batch:
            self.execute_many(sql, batch)
        print(f"  ✓ {table_name}: добавлено {count} записей")

    def run(self):
        print("🏭 Миграция НОВЫХ заказов из DBF в MySQL")
        print(f"📁 Папка: {self.path}")
        print(f"📋 Таблицы: {', '.join(TABLES)}")
        print("=" * 60)

        # Проверяем наличие всех нужных файлов
        files = {}
        for name in TABLES:
            f = self.find_dbf(name)
            if f:
                files[name] = f
            else:
                print(f"⚠️  {name}.DBF не найден — пропуск")

        if 'BASE1' not in files:
            print("❌ BASE1.DBF не найден — миграция невозможна")
            return

        print("\n📖 BASE1.DBF...")
        dbf_zaknums = self.get_dbf_zaknums(files['BASE1'])
        print(f"   Заказов в DBF: {len(dbf_zaknums)}")

        print("📖 BASE1 из MySQL...")
        mysql_zaknums = self.get_mysql_zaknums()
        print(f"   Заказов в MySQL: {len(mysql_zaknums)}")

        new_zaknums = dbf_zaknums - mysql_zaknums
        print(f"\n🆕 Новых заказов: {len(new_zaknums)}")

        if not new_zaknums:
            print("✅ Новых данных нет")
            return

        print(f"   Примеры: {', '.join(list(new_zaknums)[:5])}")

        # BASE1 первым, остальные в порядке из TABLES
        for i, name in enumerate([t for t in TABLES if t in files], 1):
            print(f"\n[{i}/{len(files)}] {name}")
            try:
                self.migrate_table(name, files[name], new_zaknums)
            except Exception as e:
                print(f"  ❌ Ошибка: {e}")

        print("\n✅ Миграция завершена")


def main():
    m = Migrator(MYSQL_CONFIG, DBF_PATH)
    try:
        m.run()
    finally:
        m.close()


if __name__ == "__main__":
    main()