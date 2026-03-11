"""
Скрипт для компиляции переводов без установки gettext
Использование: python compile_translations.py
"""
import os
import subprocess
import sys

def compile_translations():
    """Компиляция переводов используя Django команду"""
    try:
        # Пытаемся использовать стандартную команду Django
        result = subprocess.run(
            [sys.executable, 'manage.py', 'compilemessages'],
            capture_output=True,
            text=True
        )
        
        if result.returncode == 0:
            print("✓ Переводы успешно скомпилированы!")
            return True
        else:
            print("⚠ Ошибка при компиляции:")
            print(result.stderr)
            print("\n💡 Установите gettext tools или редактируйте файлы .po вручную")
            return False
    except Exception as e:
        print(f"⚠ Ошибка: {e}")
        print("\n💡 Установите gettext tools:")
        print("   choco install gettext")
        print("   или скачайте с: https://mlocati.github.io/articles/gettext-iconv-windows.html")
        return False

if __name__ == '__main__':
    print("Компиляция файлов переводов...")
    compile_translations()




