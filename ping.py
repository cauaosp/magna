import subprocess
import time

def ping_url(url, times):
    command = ["ping", url, "-n", str(times)]
    try:
        result = subprocess.run(command, capture_output=True, text=True)

        if result.returncode == 0:
            print("\nPing realizado com sucesso!")
            print(result.stdout)
        else:
            print(f"Erro ao realizar o ping: {result.stderr}")
    except Exception as e:
        print(f"Ocorreu um erro: {e}")


if __name__ == "__main__":
    while True:
        ping_url("10.197.0.60", "5")
        time.sleep(2)
