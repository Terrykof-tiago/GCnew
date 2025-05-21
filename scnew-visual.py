import os
import time
import json
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Importar serviços de drivers automáticos
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager


# Função para selecionar navegador e iniciar
def iniciar_driver(navegador="firefox"):
    if navegador.lower() == "firefox":
        options = webdriver.FirefoxOptions()
        # options.add_argument("-headless")  # Se quiser rodar sem abrir janela, descomente
        driver = webdriver.Firefox(service=FirefoxService(GeckoDriverManager().install()), options=options)

    elif navegador.lower() == "chrome":
        options = webdriver.ChromeOptions()
        # options.add_argument("-headless")
        driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    elif navegador.lower() == "edge":
        options = webdriver.EdgeOptions()
        # options.add_argument("-headless")
        driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)

    else:
        raise Exception("Navegador não suportado. Escolha 'firefox', 'chrome' ou 'edge'.")

    return driver


def executar_script():
    try:
        status_label.config(text="Processando...")

        login_url = entry_url.get()
        username = entry_user.get()
        password = entry_pass.get()
        lista_arquivo = entry_lista.get()
        navegador = navegador_var.get()

        # Verificar se arquivo da lista existe
        if not os.path.isfile(lista_arquivo):
            messagebox.showerror("Erro", "Arquivo da lista não encontrado.")
            status_label.config(text="Erro.")
            return

        # Iniciar driver com navegador selecionado
        driver = iniciar_driver(navegador)

        wait = WebDriverWait(driver, 30)

        # Login
        driver.get(login_url)
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(username)
        driver.find_element(By.NAME, "password").send_keys(password + Keys.RETURN)
        wait.until(EC.url_contains("/overview"))

        # Ler servidores
        with open(lista_arquivo, "r") as f:
            server_ids = [line.strip() for line in f if line.strip()]

        resultados = []

        for server_id in server_ids:
            api_url = f"{login_url}/api/v3.0/agents?gc_filter={server_id}&sort=display_name"

            try:
                script = f'''
                    return fetch("{api_url}")
                        .then(response => response.json())
                        .then(data => JSON.stringify(data))
                        .catch(error => "ERROR: " + error);
                '''
                result = driver.execute_script(script)

                if result.startswith("ERROR:"):
                    status = f"ERRO: {result}"
                else:
                    data = json.loads(result)
                    total = data.get("total_count", 0)
                    status = "INSTALADO" if total > 0 else "NÃO INSTALADO"

                resultados.append({"Servidor": server_id, "Status": status})
                log_text.insert(tk.END, f"{server_id}: {status}\n")
                log_text.see(tk.END)

            except Exception as e:
                resultados.append({"Servidor": server_id, "Status": f"Erro: {str(e)}"})
                log_text.insert(tk.END, f"Erro ao verificar {server_id}: {str(e)}\n")
                log_text.see(tk.END)

        driver.quit()

        # Exportar para Excel
        df = pd.DataFrame(resultados)
        output_file = "status_servidores.xlsx"
        df.to_excel(output_file, index=False)

        messagebox.showinfo("Sucesso", f"Arquivo {output_file} gerado com sucesso!")
        status_label.config(text="Concluído.")

    except Exception as e:
        messagebox.showerror("Erro", str(e))
        status_label.config(text="Erro.")


def selecionar_arquivo():
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de lista de servidores",
        filetypes=(("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*"))
    )
    if arquivo:
        entry_lista.delete(0, tk.END)
        entry_lista.insert(0, arquivo)


# Interface
janela = tk.Tk()
janela.title("Verificador de Status de Servidores")
janela.geometry("800x650")

# Campos
tk.Label(janela, text="URL de Login:").pack()
entry_url = tk.Entry(janela, width=70)
entry_url.pack()
entry_url.insert(0, "https://customer-29671397.saas.guardicore.com")

tk.Label(janela, text="Usuário:").pack()
entry_user = tk.Entry(janela, width=70)
entry_user.pack()
entry_user.insert(0, "tiago.lopes@oplium.com")

tk.Label(janela, text="Senha:").pack()
entry_pass = tk.Entry(janela, width=70, show="*")
entry_pass.pack()
entry_pass.insert(0, "M&6!r@ap1v$S2FeK")

tk.Label(janela, text="Arquivo de Lista (.txt):").pack()
frame_lista = tk.Frame(janela)
frame_lista.pack()
entry_lista = tk.Entry(frame_lista, width=55)
entry_lista.pack(side=tk.LEFT)
botao_arquivo = tk.Button(frame_lista, text="Selecionar", command=selecionar_arquivo)
botao_arquivo.pack(side=tk.LEFT)

# Seletor de navegador
tk.Label(janela, text="Selecione o navegador:").pack()
navegador_var = tk.StringVar(value="firefox")
frame_nav = tk.Frame(janela)
frame_nav.pack()

tk.Radiobutton(frame_nav, text="Firefox", variable=navegador_var, value="firefox").pack(side=tk.LEFT, padx=5)
tk.Radiobutton(frame_nav, text="Chrome", variable=navegador_var, value="chrome").pack(side=tk.LEFT, padx=5)
tk.Radiobutton(frame_nav, text="Edge", variable=navegador_var, value="edge").pack(side=tk.LEFT, padx=5)

# Log
tk.Label(janela, text="Log:").pack()
log_text = tk.Text(janela, height=20)
log_text.pack()

# Botões
tk.Button(janela, text="Executar", command=executar_script, bg="green", fg="white").pack(pady=5)

status_label = tk.Label(janela, text="Aguardando...", fg="blue")
status_label.pack()

janela.mainloop()
