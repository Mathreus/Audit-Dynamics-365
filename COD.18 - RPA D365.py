from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Credenciais de acesso
usuario = "matheus.melo@bateriasmoura.com"
senha = "@Mat108104"

# Configurações do navegador
options = webdriver.ChromeOptions()
options.add_argument("--incognito")
options.add_argument("--start-maximized")

# Inicia o navegador
driver = webdriver.Chrome(options=options)

# Acessa a página de login do D365
driver.get("https://bateriasmoura.operations.dynamics.com")

# Aguarda carregamento da página
time.sleep(3)

# Preenche o campo de usuário
campo_usuario = driver.find_element(By.ID, "i0116")
campo_usuario.send_keys(usuario)
campo_usuario.send_keys(Keys.RETURN)
time.sleep(3)

# Preenche o campo de senha
campo_senha = driver.find_element(By.ID, "i0118")
campo_senha.send_keys(senha)
campo_senha.send_keys(Keys.RETURN)
time.sleep(3)

# Clica em "Sim" para manter conectado (se aparecer)
try:
    botao_sim = driver.find_element(By.ID, "idSIButton9")
    botao_sim.click()
    time.sleep(3)
except:
    pass

# Função para executar o relatório de faturamento
def executar_faturamento():
    try:
        # Acessa o painel padrão
        driver.get("https://bateriasmoura.operations.dynamics.com/?cmp=r24&mi=DefaultDashboard")
        time.sleep(5)

        # Navega para o relatório de faturamento
        driver.get("https://bateriasmoura.operations.dynamics.com/?cmp=r24&mi=BillingStatementInquiryFilters")
        time.sleep(5)

        print("🔍 Iniciando preenchimento do relatório...")
        print(f"📄 URL atual: {driver.current_url}")
        print(f"📄 Título da página: {driver.title}")

        # PREENCHER CAMPO DA EMPRESA
        try:
            campo_empresa = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, "//input[contains(@id, 'Company') or contains(@name, 'Company')]"))
            )
            campo_empresa.clear()
            campo_empresa.send_keys("R241")
            print("✅ Campo empresa preenchido com 'R241'")
            time.sleep(3)  # Aguarda lista carregar
        except Exception as e:
            print(f"❌ Erro ao preencher campo empresa: {e}")
            return

        # CLICAR NO CHECKBOX/FLAG
        try:
            print("🔍 Procurando checkbox/flag...")
            
            # Estratégias para encontrar o checkbox
            seletores_checkbox = [
                (By.XPATH, "//div[contains(@class, 'dyn-hoverMarkingColumn')]"),
                (By.XPATH, "//div[@class='dyn-hoverMarkingColumn']"),
                (By.XPATH, "//div[contains(@class, 'public_fixedDataTableCell_cellContent')]//div[contains(@class, 'dyn-hoverMarkingColumn')]"),
                (By.XPATH, "//div[contains(@class, 'hoverMarkingColumn')]"),
                (By.XPATH, "//div[contains(@class, 'markingColumn')]"),
            ]
            
            checkbox_encontrado = False
            for seletor_type, seletor_value in seletores_checkbox:
                if not checkbox_encontrado:
                    try:
                        print(f"  Tentando encontrar checkbox com: {seletor_value}")
                        checkbox = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((seletor_type, seletor_value))
                        )
                        
                        # Verifica se o elemento está visível e clicável
                        if checkbox.is_displayed() and checkbox.is_enabled():
                            checkbox.click()
                            print(f"✅ Checkbox clicado com sucesso (usando {seletor_value})")
                            checkbox_encontrado = True
                            time.sleep(2)
                            break
                    except Exception as e:
                        print(f"  ❌ Não encontrado com {seletor_value}: {e}")
                        continue
            
            if not checkbox_encontrado:
                print("⚠️  Checkbox não encontrado com os seletores padrão. Tentando estratégia alternativa...")
                
                # Estratégia alternativa: procurar por elementos que possam ser checkboxes
                elementos_suspeitos = driver.find_elements(By.XPATH, "//div[contains(@class, 'checkbox') or contains(@class, 'check') or contains(@class, 'select') or contains(@class, 'mark')]")
                print(f"🔍 Encontrados {len(elementos_suspeitos)} elementos suspeitos de serem checkboxes")
                
                for i, elemento in enumerate(elementos_suspeitos):
                    try:
                        outer_html = elemento.get_attribute('outerHTML')
                        class_attr = elemento.get_attribute('class')
                        print(f"  Elemento {i+1}: class='{class_attr}'")
                        
                        # Se parece com um elemento de marcação/seleção, tenta clicar
                        if 'dyn-hoverMarkingColumn' in class_attr or 'hoverMarking' in class_attr or 'marking' in class_attr:
                            elemento.click()
                            print(f"✅ Checkbox clicado (elemento {i+1} com class '{class_attr}')")
                            checkbox_encontrado = True
                            time.sleep(2)
                            break
                    except Exception as e:
                        print(f"  ❌ Não foi possível clicar no elemento {i+1}: {e}")
                        continue
            
            if not checkbox_encontrado:
                print("❌ Não foi possível encontrar e clicar no checkbox. Continuando sem marcar o checkbox...")
                
        except Exception as e:
            print(f"❌ Erro ao tentar clicar no checkbox: {e}")

        # CLICAR EM SELECIONAR
        try:
            botao_selecionar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Selecionar')]"))
            )
            botao_selecionar.click()
            print("✅ Botão 'Selecionar' clicado (por qualquer elemento com texto)")
            time.sleep(3)
        except Exception as e:
            print(f"❌ Todas as tentativas falharam: {e}")
            return

        # PREENCHER DATAS - COM MÚLTIPLAS TENTATIVAS
        print("🔍 Procurando campos de data...")
        
        # Estratégias para encontrar a data inicial
        seletores_data_inicio = [
            (By.ID, "StartDate"),
            (By.XPATH, "//input[contains(@id, 'dateFrom') or contains(@name, 'dateFrom')]"),
            (By.XPATH, "//input[contains(@id, 'StartDate') or contains(@name, 'StartDate')]"),
            (By.XPATH, "//input[contains(@placeholder, 'data') or contains(@aria-label, 'data')]"),
            (By.XPATH, "//input[@type='date' or @type='text']")
        ]

        data_preenchida = False
        for seletor_type, seletor_value in seletores_data_inicio:
            if not data_preenchida:
                try:
                    data_inicio = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((seletor_type, seletor_value))
                    )
                    data_inicio.clear()
                    data_inicio.send_keys("01/09/2025")
                    print(f"✅ Data início preenchida: 01/09/2025 (usando {seletor_value})")
                    data_preenchida = True
                    break
                except:
                    continue

        if not data_preenchida:
            print("❌ Não foi possível encontrar o campo de data início")

        # Estratégias para encontrar a data final
        seletores_data_fim = [
            (By.ID, "EndDate"),
            (By.XPATH, "//input[contains(@id, 'dateTo') or contains(@name, 'dateTo')]"),
            (By.XPATH, "//input[contains(@id, 'EndDate') or contains(@name, 'EndDate')]"),
            (By.XPATH, "(//input[@type='date' or @type='text'])[2]")
        ]

        data_fim_preenchida = False
        for seletor_type, seletor_value in seletores_data_fim:
            if not data_fim_preenchida:
                try:
                    data_fim = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((seletor_type, seletor_value))
                    )
                    data_fim.clear()
                    data_fim.send_keys("30/09/2025")
                    print(f"✅ Data fim preenchida: 30/09/2025 (usando {seletor_value})")
                    data_fim_preenchida = True
                    break
                except:
                    continue

        if not data_fim_preenchida:
            print("❌ Não foi possível encontrar o campo de data fim")

        time.sleep(2)

        # DEBUG: LISTAR TODOS OS BOTÕES DISPONÍVEIS
        print("🔍 DEBUG - Listando todos os botões da página:")
        botoes = driver.find_elements(By.XPATH, "//button | //input[@type='button'] | //input[@type='submit'] | //a[@role='button'] | //div[@role='button']")
        for i, botao in enumerate(botoes):
            botao_id = botao.get_attribute("id") or ""
            botao_text = botao.text or ""
            botao_value = botao.get_attribute("value") or ""
            botao_class = botao.get_attribute("class") or ""
            if botao_id or botao_text or botao_value:
                print(f"  Botão {i+1}: id='{botao_id}', text='{botao_text}', value='{botao_value}', class='{botao_class}'")

        # CLICAR EM OK - COM MÚLTIPLAS TENTATIVAS MELHORADAS
        print("🔍 Procurando botão OK...")
        
        seletores_ok = [
            (By.ID, "billingstatementinquiryfilters_4_FormCommandButtonOK_label"),
            (By.XPATH, "//*[@id='billingstatementinquiryfilters_4_FormCommandButtonOK_label']"),
            (By.XPATH, "//button[contains(text(), 'OK')]"),
            (By.XPATH, "//input[@value='OK']"),
            (By.XPATH, "//*[text()='OK']"),
            (By.XPATH, "//button[contains(@class, 'ok') or contains(@class, 'OK')]"),
            (By.XPATH, "//input[contains(@class, 'ok') or contains(@class, 'OK')]"),
            (By.XPATH, "//a[contains(text(), 'OK')]"),
            (By.XPATH, "//div[contains(text(), 'OK')]"),
            (By.XPATH, "//span[contains(text(), 'OK')]"),
            (By.XPATH, "//*[contains(@title, 'OK') or contains(@aria-label, 'OK')]"),
            # Tentativas com números diferentes (o número pode mudar)
            (By.ID, "billingstatementinquiryfilters_3_FormCommandButtonOK_label"),
            (By.ID, "billingstatementinquiryfilters_5_FormCommandButtonOK_label"),
            (By.ID, "billingstatementinquiryfilters_6_FormCommandButtonOK_label"),
        ]

        ok_clicado = False
        for seletor_type, seletor_value in seletores_ok:
            if not ok_clicado:
                try:
                    print(f"  Tentando: {seletor_value}")
                    botao_ok = WebDriverWait(driver, 3).until(
                        EC.element_to_be_clickable((seletor_type, seletor_value))
                    )
                    botao_ok.click()
                    print(f"✅ Botão OK clicado (usando {seletor_value})")
                    ok_clicado = True
                    time.sleep(5)
                    break
                except Exception as e:
                    continue

        if not ok_clicado:
            print("❌ Não foi possível encontrar e clicar no botão OK")
            print("💡 Tente verificar manualmente qual é o ID/texto do botão OK na página")
            return

        # SELECIONAR E BAIXAR RELATÓRIO
        try:
            relatorio = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//tr[contains(@class, 'selected') or @aria-selected='true']"))
            )
            relatorio.click()
            print("✅ Relatório selecionado")

            botao_baixar = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Baixar') or contains(text(), 'Download')]"))
            )
            botao_baixar.click()
            print("✅ Botão Baixar clicado")
            
            time.sleep(8)
            print("🎉 Relatório de faturamento baixado com sucesso!")

        except Exception as e:
            print(f"❌ Erro no download: {e}")

    except Exception as e:
        print("❌ Erro geral na execução:", str(e))

# Executa a função
executar_faturamento()

# Encerra o navegador
driver.quit()