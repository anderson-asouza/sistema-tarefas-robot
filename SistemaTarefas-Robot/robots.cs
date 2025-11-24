using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System.Xml.Linq;

namespace robots
{
    public class Robo
    {
        #region Definições Gerais do Robot

        public string ERRO { get; set; } = string.Empty;
        public string DADOS { get; set; } = string.Empty;
        public bool OK { get; set; } = false;

        private static IWebDriver? _driver = null;

        public Robo(int qualwebdriver = 1)
        {
            FecharRobo();

            ERRO = "";
            DADOS = "";

            switch (qualwebdriver)
            {
                case 1:
                    _driver = new ChromeDriver();
                    break;

                case 2:
                    _driver = new FirefoxDriver();
                    break;

                default:
                    _driver = new EdgeDriver();
                    break;
            }

            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");

            //_driver.Manage().Window.Size = new System.Drawing.Size(1400, 900);
            _driver.Manage().Window.Maximize();
        }

        #endregion

        #region Funções para Robot
        private void Aguarde(int milissegundos = 1000)
        {
            if (milissegundos < 0)
                milissegundos = 1000;

            Thread.Sleep(milissegundos);
        }
        private string UrlBase(string? acessarSubPagina = null)
        {
            try
            {
                if (_driver == null)
                    return "";

                string currentUrl = _driver.Url;
                var uri = new Uri(currentUrl);

                string baseUrl = uri.IsDefaultPort
                    ? $"{uri.Scheme}://{uri.Host}"
                    : $"{uri.Scheme}://{uri.Host}:{uri.Port}";

                if (string.IsNullOrWhiteSpace(acessarSubPagina))
                {
                    return baseUrl;
                }
                else
                {
                    string url = baseUrl + "/" + acessarSubPagina.Trim();
                    _driver.Navigate().GoToUrl(url);
                    return url;
                }
            }
            catch (Exception)
            {
                return "";
            }
        }
        private void Digitar(IWebElement elemento, string texto)
        {
            try
            {
                PosiconaScrollElemento(elemento, true);
                elemento.SendKeys(OpenQA.Selenium.Keys.Control + "a");
                elemento.SendKeys(texto);
            }
            catch (Exception)
            {
            }
        }
        private bool ElementoExiste(By by, bool presenteNoFonte = false)
        {   // presenteNoFonte => Retorna True mesmo se não esteja ativo ou visível, apenas existir já basta para retornar True.
            try
            {
                if (!presenteNoFonte)
                {
                    return (_driver!.FindElements(by).Where(c => c.Displayed).Count() > 0);
                }

                return (_driver!.FindElements(by).Count > 0);
            }
            catch (Exception)
            {
                return false;
            }
        }
        private bool AguardaElementoExistir(By by, int segundos = 60, bool presenteNoFonte = false)
        {
            try
            {
                if (segundos < 1)
                {
                    segundos = 1;
                }

                DateTime horaLimite = DateTime.Now.AddSeconds(segundos);

                while (horaLimite > DateTime.Now)
                {
                    if (ElementoExiste(by, presenteNoFonte))
                    {
                        return true;
                    }

                    Aguarde(25);
                }
            }
            catch (Exception)
            {
                return false;
            }

            return false;
        }
        private int AguardaElementoDaListaExistir(List<By> lstBy, int segundos = 60, bool presenteNoFonte = false)
        {
            int r = -1;

            try
            {
                if (lstBy != null && lstBy.Count > 0)
                {
                    if (segundos < 1)
                        segundos = 1;

                    DateTime horaLimite = DateTime.Now.AddSeconds(segundos);

                    while (horaLimite > DateTime.Now)
                    {
                        Aguarde(25);

                        for (int i = 0; i < lstBy.Count; i++)
                        {
                            if (ElementoExiste(lstBy[i], presenteNoFonte))
                            {
                                r = i;
                                return r;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return r;
        }
        private bool AcessaPopUp(int indicePopUp = -1, int segundos = 10)
        {   // indicePopUp == NEGATIVO, acessa o último popup aberto. 

            try
            {
                if (segundos < 1 && indicePopUp < 0)
                {
                    segundos = 10;
                }
                else
                if (segundos < 1)
                {
                    segundos = 1;
                }

                DateTime horaLimite = DateTime.Now.AddSeconds(segundos);

                if (indicePopUp < 0)
                {
                    Aguarde(2000);

                    while (_driver!.WindowHandles.Count < 2 && horaLimite > DateTime.Now)
                    {
                        Aguarde(25);
                    }

                    if (_driver.WindowHandles.Count > 1)
                    {
                        _driver.SwitchTo().Window(_driver.WindowHandles[_driver.WindowHandles.Count - 1]);
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    while (_driver!.WindowHandles.Count <= indicePopUp)
                    {
                        if (DateTime.Now > horaLimite)
                        {
                            return false;
                        }

                        Aguarde(25);
                    }

                    _driver.SwitchTo().Window(_driver.WindowHandles[indicePopUp]);
                }

                _driver.Manage().Window.Size = new System.Drawing.Size(1360, 854);

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private bool DialogPresente(int segundos = 10)
        {
            try
            {
                if (segundos < 1)
                {
                    segundos = 1;
                }

                DateTime horaLimite = DateTime.Now.AddSeconds(segundos);

                while (DateTime.Now <= horaLimite)
                {
                    IAlert alert = new WebDriverWait(_driver!, TimeSpan.FromSeconds(10)).Until(SeleniumExtras.WaitHelpers.ExpectedConditions.AlertIsPresent());

                    if ((alert != null))
                    {
                        return true;
                    }

                    Aguarde(25);
                }

                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private bool PosiconaScrollElemento(IWebElement elemento, bool clicar = true)
        {
            try
            {
                try
                {
                    ((IJavaScriptExecutor)_driver!).ExecuteScript("arguments[0].scrollIntoView();", elemento);

                    if (clicar)
                    {
                        Aguarde(200);
                        elemento.Click();
                    }
                }
                catch (Exception)
                {
                    if (!clicar)
                    {
                        return false;
                    }

                    ((IJavaScriptExecutor)_driver!).ExecuteScript("arguments[0].click();", elemento);
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region Automações

        #region Funções para esta Automação
        private bool Limpar()
        {
            DADOS = string.Empty;
            ERRO = string.Empty;
            OK = false;

            if (_driver == null)
            {
                ERRO = "Robot não foi iniciado!";
                return false;
            }

            return true;
        }
        private bool EstaLogadoOUAbreLogin()
        {
            By elemMenu = By.TagName("nav");

            if (!AguardaElementoExistir(elemMenu, 30))
            {
                ERRO = "Menu não foi encontrado.";
                return false;
            }

            var lis = _driver!.FindElement(elemMenu).FindElements(By.TagName("li")).Where(c => c.Displayed);

            if (lis.Count() > 1)
                return true;

            var menuLogin = lis.Where(c => c.Text.ToUpper().Trim() == "LOGIN").ToList();

            if (menuLogin.Count != 1)
            {
                ERRO = "Menu de Login não foi encontrado.";
                return false;
            }

            PosiconaScrollElemento(menuLogin[0], true);

            Aguarde(200);

            if (menuLogin[0].FindElements(By.TagName("div")).Count <= 2)
                return false;

            PosiconaScrollElemento(menuLogin[0], true);
            return true;
        }
        private bool AcessaMenu(string mainMenu, string itemMenu)
        {
            if (!EstaLogadoOUAbreLogin())
            {
                ERRO = "Não está logado.";
                return false;
            }

            By elemMenu = By.TagName("nav");

            if (!AguardaElementoExistir(elemMenu, 30))
            {
                ERRO = "Menu não foi encontrado.";
                return false;
            }

            var lis = _driver!.FindElement(elemMenu).FindElements(By.TagName("li")).Where(c => c.Displayed);

            if (lis.Count() == 1)
            {
                ERRO = "Perfil sem opções de menu. Possivelmente usuário bloquado.";
                return false;
            }

            var menu = lis.Where(c => c.Text.ToUpper().Trim() == mainMenu.ToUpper().Trim()).ToList();

            if (menu.Count != 1)
            {
                ERRO = $"Menu '{mainMenu.Trim()}' não está disponível para este perfil.";
                return false;
            }

            PosiconaScrollElemento(menu[0], true);

            Aguarde(250);
            var item = menu[0].FindElements(By.TagName("button")).Where(c => c.Text.ToUpper().Trim() == itemMenu.ToUpper().Trim()).ToList();

            if (item.Count != 1)
            {
                PosiconaScrollElemento(menu[0], true);
                ERRO = $"Item de Menu '{itemMenu.Trim()}' não está disponível para este perfil.";
                return false;
            }

            PosiconaScrollElemento(item[0], true);

            return true;
        }

        private bool AguardaLoading(By? elem = null, int segundos = 60, int aguardeIniciarMilissegundos = 250) {
            try
            {
                if (segundos < 1)
                    segundos = 1;

                if (aguardeIniciarMilissegundos < 0)
                    aguardeIniciarMilissegundos = 0;

                if (elem == null)
                    elem = By.Id("loading");

                Aguarde(aguardeIniciarMilissegundos);
                DateTime horaLimite = DateTime.Now.AddSeconds(segundos);

                while (ElementoExiste(elem))
                {
                    if (DateTime.Now > horaLimite)
                    {
                        ERRO = "Tempo limite de espera no Loading foi excedido.";
                        return false;
                    }
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region Início
        public void FecharRobo()
        {
            if (_driver != null)
            {
                _driver.Quit();
                _driver = null;
            }
        }
        public void AbrirWebSite(string url)
        {

            try
            {
                if (!Limpar())
                    return;

                _driver!.Navigate().GoToUrl(url);

                DADOS = "Robot Iniciado.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #region Menu Login
        public void EfetuarLogin(string login, string senha, string novaSenha = "", string confirmacaoNovaSenha = "")
        {
            try
            {
                if (!Limpar())
                    return;

                By elemLogin = By.Id("login");
                By elemSenha = By.Id("senha");
                By elemNovaSenha = By.Id("novaSenha");
                By elemConfirmacaoNovaSenha = By.Id("confirmacaoNovaSenha");

                if (EstaLogadoOUAbreLogin())
                {
                    ERRO = "Já está logado.";
                    OK = true;
                    return;
                }

                if (ERRO != "")
                    return;

                if (!AguardaElementoExistir(elemLogin, 10))
                {
                    ERRO = "Elementos de login não carregaram.";
                    return;
                }

                bool existeCampoNovaSenha = ElementoExiste(elemNovaSenha);

                if (existeCampoNovaSenha && novaSenha == "")
                {
                    ERRO = "Este login está solicitando alteração de senha agora. Informe uma nova senha.";
                    return;
                }

                Digitar(_driver!.FindElement(elemLogin), login);
                Digitar(_driver.FindElement(elemSenha), senha);

                if (existeCampoNovaSenha)
                {
                    Digitar(_driver.FindElement(elemNovaSenha), novaSenha);
                    Digitar(_driver.FindElement(elemConfirmacaoNovaSenha), confirmacaoNovaSenha);
                }

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Where(c => c.Text.ToUpper().Trim() == "LOGAR").ToList()[0], true);

                DateTime horaLimite = DateTime.Now.AddSeconds(60);

                while (!EstaLogadoOUAbreLogin())
                {
                    if (ElementoExiste(By.ClassName("custom-alert-message")))
                    {
                        ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                        return;
                    }

                    if (DateTime.Now > horaLimite)
                    {
                        ERRO = "Login não foi efetuado.";
                        return;
                    }
                }

                DADOS = "Login Efetuado com sucesso.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void AlterarSenha(string login, string senha, string novaSenha, string confirmacaoNovaSenha)
        {
            try
            {
                if (!Limpar())
                    return;

                By elemLogin = By.Id("login");
                By elemSenha = By.Id("senha");
                By elemNovaSenha = By.Id("novaSenha");
                By elemConfirmacaoNovaSenha = By.Id("confirmacaoNovaSenha");

                if (_driver!.FindElements(By.TagName("h1")).Where(c => c.Text.ToUpper().Contains("ALTERAÇÃO DE SENHA")).Count() == 0)
                {

                    if (!EstaLogadoOUAbreLogin())
                    {
                        ERRO = "Não está logado. Efetue o Login.";
                        return;
                    }

                    if (!AcessaMenu("Login", "Alterar Senha"))
                        return;
                }

                if (!AguardaElementoExistir(elemLogin, 10))
                {
                    ERRO = "Elementos para alterar a senha não carregaram.";
                    return;
                }

                Digitar(_driver!.FindElement(elemLogin), login);
                Digitar(_driver.FindElement(elemSenha), senha);
                Digitar(_driver.FindElement(elemNovaSenha), novaSenha);
                Digitar(_driver.FindElement(elemConfirmacaoNovaSenha), confirmacaoNovaSenha);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "ALTERAR"), true);

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação do Cadastro do Modelo de Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "OK"), true);
                    return;
                }

                DADOS = "Senha do usuário foi alterada com sucesso.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void MudarUsuario(string login, string senha)
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o login.";
                    return;
                }

                if (!AcessaMenu("Login", "Mudar de Usuário"))
                {
                    ERRO = "Falha ao acessar o menu.";
                    return;
                }

                EfetuarLogin(login, senha);
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void AdicionarFoto(string caminho)
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o login.";
                    return;
                }
               
                IJavaScriptExecutor js = (IJavaScriptExecutor)_driver!;

                js.ExecuteScript("document.getElementById('imagemInput').value = '';");

                IWebElement fileInput = _driver!.FindElement(By.Id("imagemInput"));
                fileInput.SendKeys(caminho);

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação do envio da foto.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "OK"), true);
                    return;
                }

                DADOS = "Foto de Perfil foi adicionada com sucesso!";
                OK = true;
            }
            catch (Exception ex)
            {
                try
                {
                    IJavaScriptExecutor js = (IJavaScriptExecutor)_driver!;
                    js.ExecuteScript("document.getElementById('imagemInput').style.display = 'none';");
                }
                catch (Exception)
                {
                }

                ERRO = ex.Message;
                return;
            }
        }
        public void RemoverFoto()
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o login.";
                    return;
                }

                if (!AcessaMenu("Login", "Remover Foto"))
                {
                    ERRO = "Usuário não possui foto.";
                    OK = true;
                    return;
                }

                if (!AguardaElementoExistir(By.ClassName("custom-alert-message")))
                {
                    ERRO = "Dialog de confirmação 'Remover Foto' não foi identificado.";
                    return;
                }

                Aguarde(250);
                var sim = _driver!.FindElements(By.TagName("button")).Where(c => c.Text.ToUpper().Trim() == "SIM").ToList();

                if (sim.Count != 1)
                {
                    ERRO = "Dialog não abriu corretamente.";
                    return;
                }
                sim[0].Click();

                DADOS = "Foto de Perfil removida.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void EfetuarLogoff()
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Já está em Logoff. Usuário não está logado.";
                    OK = true;
                    return;
                }

                if (!AcessaMenu("Login", "Fazer Logoff"))
                    return;

                if (!AguardaElementoExistir(By.ClassName("custom-alert-message")))
                {
                    ERRO = "Dialog de confirmação do Logoff não foi identificado.";
                    return;
                }

                Aguarde(250);
                var sim = _driver!.FindElements(By.TagName("button")).Where(c => c.Text.ToUpper().Trim() == "SIM").ToList();

                if (sim.Count != 1)
                {
                    ERRO = "Dialog não abriu corretamente.";
                    return;
                }
                sim[0].Click();

                DADOS = "O Logoff foi efetuado.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #region Usuários
        public void CadastroUsuario(bool cadastroNovo, string nomeConsulta, string nomeCadastro, string login, string senha, string nivel, string email, string matricula)
        {
            try
            {
                if (!cadastroNovo)
                {
                    ConsultarUsuario(nomeConsulta);

                    if (!OK)
                    {
                        ERRO = "Atualização - " + ERRO;
                        return;
                    }

                    PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[1], true);
                }
                else
                {
                    if (!Limpar())
                        return;

                    if (!AcessaMenu("Usuários", "Cadastrar Usuário"))
                        return;
                }

                By elemLogin = By.Id("login");
                By elemSenha = By.Id("senha");
                By elemNome = By.Id("nome");
                By elemNivel = By.Id("nivel");
                By elemEmail = By.Id("email");
                By elemMatricula = By.Id("matricula");

                if (!AguardaElementoExistir(elemLogin, 10))
                {
                    ERRO = "Cadastro do Usuário não carregou corretamente.";
                    return;
                }

                if (cadastroNovo)
                {
                    Digitar(_driver!.FindElement(elemLogin), login);
                    Digitar(_driver.FindElement(elemSenha), senha);
                }

                Digitar(_driver!.FindElement(elemNivel), nivel);
                Digitar(_driver.FindElement(elemNome), nomeCadastro);
                Digitar(_driver.FindElement(elemEmail), email);
                Digitar(_driver.FindElement(elemMatricula), matricula);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "CADASTRAR" || c.Text.ToUpper().Trim() == "ATUALIZAR"), true);

                AguardaElementoExistir(By.ClassName("custom-alert-message"));
                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM"), true);

                Aguarde(200);
                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = $"Não houve confirmação {(cadastroNovo ? "do Cadastro" : "da Atualização")} do Usuário.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    return;
                }

                DADOS = $"Usuário {(cadastroNovo ? "Cadastrado" : "Atualizado")} com sucesso.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ConsultarUsuario(string nome)
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o Login.";
                    return;
                }

                UrlBase("usuarios");

                List<By> elems = new List<By>();
                elems.Add(By.Id("nomeUsuario"));

                int i = AguardaElementoDaListaExistir(elems, 10);

                if (i < 0)
                {
                    ERRO = "Usuário não tem acesso a Gerência de Usuários.";
                    return;
                }

                if (!AguardaLoading())
                    return;

                Digitar(_driver!.FindElement(elems[0]), nome);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "LISTAR"), true);

                if (!AguardaLoading())
                    return;

                var item = _driver!.FindElement(By.TagName("table")).FindElement(By.TagName("tbody")).FindElements(By.TagName("tr")).ToList();

                if (item.Count() == 0)
                {
                    ERRO = "Registro consultado não foi encontado.";
                    return;
                }

                if (item.Count() > 1)
                {
                    ERRO = "Deve ser mais específio para poder identificar o Usuário corretamente.";
                    return;
                }

                PosiconaScrollElemento(item[0], true);

                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ExcluirUsuario(string nome)
        {
            try
            {
                ConsultarUsuario(nome);

                if (!OK || ERRO != "")
                {
                    ERRO = "Exclusão - " + ERRO;
                    return;
                }

                OK = false;
                DADOS = "";

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[2], true);

                AguardaElementoExistir(elems[1]);

                _driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();
                Aguarde(200);

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação da Exclusão do usuário.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(elems[1]).Text;
                    return;
                }

                DADOS = "Usuário foi excluído.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void AlterarSenhaPeloAdm(string login, string novaSenha, string confirmacaoNovaSenha)
        {
            try
            {
                if (!Limpar())
                    return;

                By elemLogin = By.Id("login");
                By elemNovaSenha = By.Id("novaSenha");
                By elemConfirmacaoNovaSenha = By.Id("confirmacaoNovaSenha");

                if (_driver!.FindElements(By.TagName("h1")).Where(c => c.Text.ToUpper().Contains("ALTERAÇÃO DE SENHA")).Count() == 0)
                {

                    if (!EstaLogadoOUAbreLogin())
                    {
                        ERRO = "Não está logado. Efetue o Login.";
                        return;
                    }

                    if (!AcessaMenu("Usuários", "Alterar Senha como Administrador"))
                        return;
                }

                if (!AguardaElementoExistir(elemLogin, 10))
                {
                    ERRO = "Elementos para alterar a senha não carregaram.";
                    return;
                }

                Digitar(_driver!.FindElement(elemLogin), login);
                Digitar(_driver.FindElement(elemNovaSenha), novaSenha);
                Digitar(_driver.FindElement(elemConfirmacaoNovaSenha), confirmacaoNovaSenha);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "ALTERAR"), true);

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação do Cadastro do Modelo de Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "OK"), true);

                    return;
                }

                DADOS = "Senha do usuário foi alterada com sucesso.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #region Modelos Tarefa
        public void CadastroModeloTarefa(bool cadastroNovo, string nomeConsulta, string nome, string descricao)
        {
            try
            {
                if (!Limpar())
                    return;

                if (cadastroNovo)
                {
                    if (_driver!.FindElements(By.TagName("h1")).Where(c => c.Text.ToUpper().Trim() == "CADASTRO DE MODELO DE TAREFA").Count() == 0)
                    {
                        if (!EstaLogadoOUAbreLogin())
                        {
                            ERRO = "Usuário não está logado. Efetue o login.";
                            return;
                        }

                        if (!AcessaMenu("Modelos", "Cadastrar Modelo de Tarefa"))
                        {
                            ERRO = "Usuário não tem acesso ao Modelo de Tarefa.";
                            return;
                        }
                    }
                }
                else
                {
                    ConsultarModeloTarefa(nomeConsulta);

                    if (!OK || ERRO != "")
                    {
                        ERRO = "Atualização - " + ERRO;
                        return;
                    }

                    PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[1], true);
                }

                By elemNome = By.Id("nomeTarefa");
                By elemDescricao = By.Id("descricaoTarefa");

                if (!AguardaElementoExistir(elemNome, 10))
                {
                    ERRO = "Cadastro do Modelo de Tarefas não carregou corretamente.";
                    return;
                }

                Digitar(_driver!.FindElement(elemNome), nome);
                Digitar(_driver.FindElement(elemDescricao), descricao);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "CADASTRAR" || c.Text.ToUpper().Trim() == "ATUALIZAR"), true);

                AguardaElementoExistir(By.ClassName("custom-alert-message"));

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM"), true);
                Aguarde(200);

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação do Cadastro do Modelo de Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "OK"), true);

                    return;
                }

                DADOS = "Modelo de Tarefa foi gravado com sucesso.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ConsultarModeloTarefa(string nome)
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o Login.";
                    return;
                }

                UrlBase("modelostarefa");

                List<By> elems = new List<By>();
                elems.Add(By.Id("nomeModeloTarefa"));
                elems.Add(By.ClassName("custom-alert-message"));
                elems.Add(By.TagName("table"));

                if (!AguardaElementoExistir(elems[0], 10))
                {
                    ERRO = "Usuário não tem acesso ao Modelo de Tarefa.";
                    return;
                }

                if (ElementoExiste(elems[1]))
                {
                    ERRO = _driver!.FindElement(elems[1]).Text;
                    return;
                }

                if (!AguardaLoading())
                    return;

                if (!ElementoExiste(elems[2]))
                {
                    ERRO = "Nenhum registro encontado.";
                    OK = true;
                    return;
                }

                Digitar(_driver!.FindElement(elems[0]), nome);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "LISTAR"), true);

                if (!AguardaLoading())
                    return;

                if (_driver.FindElement(elems[2]).FindElement(By.TagName("tbody")).FindElements(By.TagName("tr")).Count() == 0)
                {
                    ERRO = "Registro consultado não foi encontado.";
                    return;
                }

                if (_driver.FindElement(elems[2]).FindElement(By.TagName("tbody")).FindElements(By.TagName("tr")).Count() > 1)
                {
                    ERRO = "Informe um nome mais específio para identificar o Modelo de Tarefa corretamente.";
                    return;
                }

                PosiconaScrollElemento(_driver.FindElement(elems[2]).FindElement(By.TagName("tbody")).FindElement(By.TagName("tr")), true);

                DADOS = "Modelo de Tarefa foi encontrado com sucesso.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ExcluirModeloTarefa(string nome)
        {
            try
            {
                ConsultarModeloTarefa(nome);

                if (!OK || ERRO != "")
                {
                    ERRO = "Exclusão - " + ERRO;
                    return;
                }

                OK = false;
                DADOS = "";

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[2], true);

                AguardaElementoExistir(elems[1]);

                _driver!.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();
                Aguarde(200);

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação da Exclusão do Modelo de Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(elems[1]).Text;
                    return;
                }

                DADOS = "Cadastro foi excluído.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #region Trâmites Modelo Tarefa
        public void CadastroTramiteModeloTarefa(bool cadastroNovo, string vinculo, string nomeConsulta, string nome, string descricao, string duracao, string revisor, string tramitador)
        {
            try
            {
                ConsultarTramiteModeloTarefa(vinculo, nomeConsulta);

                if (!OK)
                {
                    ERRO = cadastroNovo ? "Cadastro - " : "Atualização - " + ERRO;
                    return;
                }

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[(cadastroNovo) ? 1 : 2], true);

                By elemNome = By.Id("nomeTramite");
                By elemDescricao = By.Id("descricaoTramite");
                By elemDuracao = By.Id("duracaoPrevistaDias");
                By elemRevisor = By.Id("usuarioRevisor");
                By elemTramitador = By.Id("usuarioIndicacao");

                if (!AguardaElementoExistir(elemNome, 10))
                {
                    ERRO = "Cadastro do Trâmite Modelo de Tarefas não carregou corretamente.";
                    return;
                }

                Digitar(_driver.FindElement(elemNome), nome);
                Digitar(_driver.FindElement(elemDescricao), descricao);
                Digitar(_driver.FindElement(elemDuracao), duracao);

                if (!string.IsNullOrWhiteSpace(revisor))
                {
                    PosiconaScrollElemento(_driver.FindElement(elemRevisor), true);
                    Aguarde(200);

                    var item = _driver.FindElement(elemRevisor).FindElements(By.TagName("option")).Where(c => c.Text.ToUpper().Contains(revisor.ToUpper().Trim())).ToList();

                    if (item.Count() == 0)
                    {
                        ERRO = "Nome de usuário para Revisor não foi encontrado.";
                        return;
                    }

                    if (item.Count() > 1)
                    {
                        ERRO = "Informe um nome mais específio para definir o Usuário Revisor.";
                        return;
                    }

                    PosiconaScrollElemento(item[0], true);
                }

                if (!string.IsNullOrWhiteSpace(tramitador))
                {
                    PosiconaScrollElemento(_driver.FindElement(elemTramitador), true);
                    Aguarde(200);

                    var item = _driver.FindElement(elemTramitador).FindElements(By.TagName("option")).Where(c => c.Text.ToUpper().Contains(tramitador.ToUpper().Trim())).ToList();

                    if (item.Count() == 0)
                    {
                        ERRO = "Nome de usuário para Tramitador não foi encontrado.";
                        return;
                    }

                    if (item.Count() > 1)
                    {
                        ERRO = "Informe um nome mais específio para definir o Usuário Tramitador.";
                        return;
                    }

                    PosiconaScrollElemento(item[0], true);
                }

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "CADASTRAR" || c.Text.ToUpper().Trim() == "ATUALIZAR"), true);

                AguardaElementoExistir(By.ClassName("custom-alert-message"));
                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM"), true);

                Aguarde(200);
                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = $"Não houve confirmação {(cadastroNovo ? "do Cadastro" : "da Atualização")} para o Trâmite do Modelo de Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    return;
                }

                DADOS = $"Trâmite para Modelo de Tarefa {(cadastroNovo ? "Cadastrado" : "Atualizado")} com sucesso.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ConsultarTramiteModeloTarefa(string nomeModeloTarefa, string nomeModeloTramite)
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o Login.";
                    return;
                }

                UrlBase("modelostramite");

                By elemTarefa = By.Id("tarefas");

                if (!AguardaElementoExistir(elemTarefa, 10))
                {
                    ERRO = "Usuário não tem acesso ao Trâmites para Modelo de Tarefa.";
                    return;
                }

                if (!AguardaLoading())
                    return;

                PosiconaScrollElemento(_driver!.FindElement(elemTarefa), true);
                Aguarde(200);

                var modeloTarefa = _driver!.FindElement(elemTarefa).FindElements(By.TagName("option")).Where(c => c.Text.ToUpper().Contains(nomeModeloTarefa.ToUpper().Trim())).ToList();

                if (modeloTarefa.Count() == 0)
                {
                    ERRO = "Registro consultado não foi encontado.";
                    return;
                }

                if (modeloTarefa.Count() > 1)
                {
                    ERRO = "Informe um nome mais específio para identificar o Modelo de Tarefa corretamente.";
                    return;
                }

                PosiconaScrollElemento(modeloTarefa[0], true);

                if (!AguardaLoading())
                    return;

                if (!string.IsNullOrWhiteSpace(nomeModeloTramite))
                {
                    if (!AguardaElementoExistir(By.TagName("table"), 10))                    
                    {
                        ERRO = "Não foi encontrado nenhum registro.";
                        return;
                    }

                    var item = _driver!.FindElement(By.TagName("table")).FindElement(By.TagName("tbody")).FindElements(By.TagName("tr")).Where(c => c.FindElements(By.TagName("td"))[1].Text.ToUpper().Contains(nomeModeloTramite.ToUpper().Trim())).ToList();

                    if (item.Count() == 0)
                    {
                        ERRO = "Registro não foi encontrado";
                        return;
                    }

                    if (item.Count() > 1)
                    {
                        ERRO = "Informe um nome mais específio para identificar o Modelo do Trâmite corretamente.";
                        return;
                    }

                    PosiconaScrollElemento(item[0], true);
                }

                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ExcluirTramiteModeloTarefa(string vinculo, string nome)
        {
            try
            {
                ConsultarTramiteModeloTarefa(vinculo, nome);

                if (!OK || ERRO != "")
                {
                    ERRO = "Exclusão - " + ERRO;
                    return;
                }

                OK = false;
                DADOS = "";

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[3], true);

                AguardaElementoExistir(elems[1]);

                _driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();
                Aguarde(200);

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação da Exclusão do Modelo de Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(elems[1]).Text;
                    return;
                }

                DADOS = "Cadastro foi excluído.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #region Flags
        public void CadastroFlag(bool cadastroNovo, string nomeConsulta, string nomeCadastro, string cor)
        {
            try
            {
                if (!cadastroNovo)
                {
                    ConsultarFlag(nomeConsulta);

                    if (!OK)
                    {
                        ERRO = "Atualização - " + ERRO;
                        return;
                    }

                    PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[1], true);
                }
                else
                {
                    if (!Limpar())
                        return;

                    if (!AcessaMenu("Tarefas", "Cadastrar Flag"))
                        return;
                }

                By elemNome = By.Id("nome");
                By elemCor = By.Id("cor");

                if (!AguardaElementoExistir(elemNome, 10))
                {
                    ERRO = "Cadastro da Flag não carregou corretamente.";
                    return;
                }

                Digitar(_driver!.FindElement(elemNome), nomeCadastro);
                Digitar(_driver.FindElement(elemCor), cor);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "CADASTRAR" || c.Text.ToUpper().Trim() == "ATUALIZAR"), true);

                AguardaElementoExistir(By.ClassName("custom-alert-message"));
                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM"), true);

                Aguarde(200);
                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = $"Não houve confirmação {(cadastroNovo ? "do Cadastro" : "da Atualização")} da Flag.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    return;
                }

                DADOS = $"Flag {(cadastroNovo ? "Cadastrada" : "Atualizada")} com sucesso.";
                OK = true;

                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ConsultarFlag(string nome)
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o Login.";
                    return;
                }

                UrlBase("flags");

                List<By> elems = new List<By>();
                elems.Add(By.Id("RotuloFlag"));
                elems.Add(By.ClassName("custom-alert-message"));                

                if (!AguardaElementoExistir(elems[0], 10))
                {
                    ERRO = "Usuário não tem acesso a Gerência de Flags para Tarefas.";
                    return;
                }

                if (ElementoExiste(elems[1]))
                {
                    ERRO = _driver!.FindElement(elems[1]).Text;
                    return;
                }

                if (!AguardaLoading())
                    return;

                Digitar(_driver!.FindElement(elems[0]), nome);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "LISTAR"), true);

                if (!AguardaLoading())
                    return;

                var item = _driver!.FindElement(By.TagName("table")).FindElement(By.TagName("tbody")).FindElements(By.TagName("tr")).ToList();

                if (item.Count() == 0)
                {
                    ERRO = "Registro consultado não foi encontado.";
                    return;
                }

                if (item.Count() > 1)
                {
                    ERRO = "Informe um nome mais específio para identificar a Flag corretamente.";
                    return;
                }

                PosiconaScrollElemento(item[0], true);

                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ExcluirFlag(string nome)
        {
            try
            {
                ConsultarFlag(nome);

                if (!OK || ERRO != "")
                {
                    ERRO = "Exclusão - " + ERRO;
                    return;
                }

                OK = false;
                DADOS = "";

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[2], true);

                AguardaElementoExistir(elems[1]);

                _driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();
                Aguarde(200);

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação da Exclusão da Flag.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(elems[1]).Text;
                    return;
                }

                DADOS = "Flag foi excluída.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #region Tarefas
        public void CadastroTarefa(bool cadastroNovo, string vinculo, string nomeConsulta, string nomeCadastro, string descricao)
        {
            try
            {
                if (!cadastroNovo)
                {
                    ConsultarTarefa(nomeConsulta);

                    if (!OK)
                    {
                        ERRO = "Atualização - " + ERRO;
                        return;
                    }

                    PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[1], true);
                }
                else
                {
                    if (!Limpar())
                        return;

                    if (!AcessaMenu("Tarefas", "Cadastrar Tarefa"))
                        return;
                }

                By elemViculo = By.Id("modeloTarefas");
                By elemNome = By.Id("nomeTarefa");
                By elemDescricao = By.Id("descricaoTarefa");

                if (!AguardaElementoExistir(elemNome, 10))
                {
                    ERRO = "Cadastro da Tarefa não carregou corretamente.";
                    return;
                }

                if (cadastroNovo)
                {
                    PosiconaScrollElemento(_driver!.FindElement(elemViculo), true);
                    Aguarde(200);

                    var item = _driver.FindElement(elemViculo).FindElements(By.TagName("option")).Where(c => c.Text.ToUpper().Contains(vinculo.ToUpper().Trim())).ToList();

                    if (item.Count() == 0)
                    {
                        ERRO = "Nome do Modelo de Tarefa não foi encontrado.";
                        return;
                    }

                    if (item.Count() > 1)
                    {
                        ERRO = "Informe um nome mais específio para definir o Modelo de Tarefa corretamente.";
                        return;
                    }

                    PosiconaScrollElemento(item[0], true);
                }

                Digitar(_driver!.FindElement(elemNome), nomeCadastro);
                Digitar(_driver.FindElement(elemDescricao), descricao);

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "CADASTRAR" || c.Text.ToUpper().Trim() == "ATUALIZAR"), true);

                AguardaElementoExistir(By.ClassName("custom-alert-message"));
                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM"), true);

                Aguarde(200);
                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = $"Não houve confirmação {(cadastroNovo ? "do Cadastro" : "da Atualização")} da Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    return;
                }

                DADOS = $"Tarefa {(cadastroNovo ? "Cadastrada" : "Atualizada")} com sucesso.";
                OK = true;

                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ConsultarTarefa(string nome, string tipoStatus = "Abertas")
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o Login.";
                    return;
                }

                UrlBase("tarefas");

                List<By> elems = new List<By>();
                elems.Add(By.Id("tarefa"));
                elems.Add(By.Id("statusTarefa"));
                elems.Add(By.ClassName("custom-alert-message"));

                if (!AguardaElementoExistir(elems[0], 10))
                {
                    ERRO = "Usuário não tem acesso a Gerência das Tarefas.";
                    return;
                }

                if (ElementoExiste(elems[2]))
                {
                    ERRO = _driver!.FindElement(elems[2]).Text;
                    return;
                }

                if (!AguardaLoading())
                    return;

                Digitar(_driver!.FindElement(elems[0]), nome);

                if (tipoStatus != "Abertas")
                {
                    PosiconaScrollElemento(_driver.FindElement(elems[1]).FindElements(By.TagName("option")).Single(c => c.Text.ToUpper().Contains(tipoStatus.ToUpper().Trim())), true);
                }

                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "LISTAR"), true);

                if (!AguardaLoading())
                    return;

                var item = _driver!.FindElement(By.TagName("table")).FindElement(By.TagName("tbody")).FindElements(By.TagName("tr")).ToList();

                if (item.Count() == 0)
                {
                    ERRO = "Registro consultado não foi encontado.";
                    return;
                }

                if (item.Count() > 1)
                {
                    ERRO = "Informe um nome mais específio para identificar a Tarefa corretamente.";
                    return;
                }

                PosiconaScrollElemento(item[0], true);

                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void AtivarDesativarTarefa(string nome, bool ativar = true)
        {
            try
            {
                ConsultarTarefa(nome, (ativar) ? "Desativadas" : "Abertas");

                if (!OK || ERRO != "")
                {
                    ERRO = (ativar ? "Ativar" : "Desativar") +" - "+ ERRO +(ativar ? " Para Ativação." : " Para Desativação.");
                    return;
                }

                OK = false;
                DADOS = "";

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[2], true);

                AguardaElementoExistir(elems[1]);

                _driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();
                Aguarde(200);

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação da Exclusão da Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(elems[1]).Text;
                    return;
                }

                DADOS = "Tarefa foi "+(ativar ? "Ativada com Sucesso." : "Desativada com sucesso.");
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ExcluirTarefa(string nome)
        {
            try
            {
                ConsultarTarefa(nome, "Desativadas");

                if (!OK || ERRO != "")
                {
                    ERRO = "Exclusão - " + ERRO+ " A tarefa precisa estar desativada para ser excluída.";
                    return;
                }

                OK = false;
                DADOS = "";

                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[3], true);

                AguardaElementoExistir(elems[1]);

                _driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();
                Aguarde(200);

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Não houve confirmação da Exclusão da Tarefa.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(elems[1]).Text;
                    return;
                }

                DADOS = "Tarefa foi excluída.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #region Trâmites

        public void IncluirTramite(string tarefa)
        {
            try
            {
                ConsultarTramite(tarefa);

                if (!OK)
                {
                    ERRO = "Incluir Trâmite - " + ERRO;
                    return;
                }

                DADOS = "";
                OK = false;

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[1], true);

                AguardaElementoExistir(By.ClassName("custom-alert-message"));
                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM"), true);

                Aguarde(200);
                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = $"Não houve confirmação inclusão de um Trâmite.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    return;
                }

                DADOS = $"Trâmite incluso de forma manual.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ConsultarTramite(string nomeTarefa, string nomeTramite = "")
        {
            try
            {
                if (!Limpar())
                    return;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o Login.";
                    return;
                }

                UrlBase("tramites");

                By elemTarefa = By.Id("tarefas");

                if (!AguardaElementoExistir(elemTarefa, 10))
                {
                    ERRO = "Usuário não tem acesso aos Trâmites.";
                    return;
                }

                if (!AguardaLoading())
                    return;

                if (ElementoExiste(By.ClassName("custom-alert-message")))
                {
                    var ok = _driver!.FindElements(By.TagName("button")).Where(c => c.Text.ToUpper().Trim() == "OK").ToList();
                    if (ok.Count > 0)
                        PosiconaScrollElemento(ok[0], true);
                }

                PosiconaScrollElemento(_driver!.FindElement(elemTarefa), true);
                Aguarde(200);

                var tarefaSelecionada = _driver!.FindElement(elemTarefa).FindElements(By.TagName("option")).Where(c => c.Text.ToUpper().Contains(nomeTarefa.ToUpper().Trim())).ToList();

                if (tarefaSelecionada.Count() == 0)
                {
                    ERRO = "Registro consultado não foi encontado.";
                    return;
                }

                if (tarefaSelecionada.Count() > 1)
                {
                    ERRO = "Informe um nome mais específio para identificar a Tarefa corretamente.";
                    return;
                }

                PosiconaScrollElemento(tarefaSelecionada[0], true);

                if (!AguardaLoading())
                    return;

                if (!string.IsNullOrWhiteSpace(nomeTramite))
                {
                    if (!AguardaElementoExistir(By.TagName("table"), 10))
                    {
                        ERRO = "Não foi encontrado nenhum registro.";
                        return;
                    }

                    var item = _driver!.FindElement(By.TagName("table")).FindElement(By.TagName("tbody")).FindElements(By.TagName("tr")).Where(c => c.FindElements(By.TagName("td"))[2].Text.ToUpper().Contains(nomeTramite.ToUpper().Trim())).ToList();

                    if (item.Count() == 0)
                    {
                        ERRO = "Registro não foi encontrado";
                        return;
                    }

                    if (item.Count() > 1)
                    {
                        ERRO = "Informe um nome mais específio para selecionar o Trâmite corretamente.";
                        return;
                    }

                    PosiconaScrollElemento(item[0], true);
                }

                DADOS = "Trâmite encontrado.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void VoltarTramite(string tarefa, string tramite)
        {
            try
            {
                ConsultarTramite(tarefa, tramite);

                if (!OK)
                {
                    ERRO = "Voltar Trâmite - " + ERRO;
                    return;
                }

                DADOS = "";
                OK = false;

                PosiconaScrollElemento(_driver!.FindElements(By.ClassName("colunas")).ToList()[0].FindElements(By.TagName("svg")).ToList()[2], true);

                AguardaElementoExistir(By.ClassName("custom-alert-message"));
                PosiconaScrollElemento(_driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM"), true);

                Aguarde(200);
                List<By> elems = new List<By>();
                elems.Add(By.ClassName("splash-text"));
                elems.Add(By.ClassName("custom-alert-message"));

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = $"Não houve confirmação de voltar Trâmite.";
                    return;
                }

                if (i == 1)
                {
                    ERRO = _driver.FindElement(By.ClassName("custom-alert-message")).Text;
                    return;
                }

                if (AguardaElementoExistir(elems[1], 2))
                {
                    var ok = _driver.FindElements(By.TagName("button")).Where(c => c.Text.ToUpper().Trim() == "OK").ToList();
                    if (ok.Count > 0)
                        PosiconaScrollElemento(ok[0], true);
                }

                DADOS = $"Trâmite retroagido de forma manual.";
                OK = true;
                return;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #region Cards

         public IWebElement? ConsultarCard(string nomeTarefa, bool abrir = false)
        {
            try
            {
                if (!Limpar())
                    return null;

                if (!EstaLogadoOUAbreLogin())
                {
                    ERRO = "Usuário não está logado. Efetue o Login.";
                    return null;
                }

                UrlBase("cards");

                if (!AguardaLoading())
                    return null;

                List<By> elems = new List<By>();
                elems.Add(By.XPath("//h1[contains(., 'Não há cards para este usuário')]"));
                elems.Add(By.XPath("//div[contains(@class, 'listaCards')]"));
                

                int i = AguardaElementoDaListaExistir(elems);

                if (i < 0)
                {
                    ERRO = "Usuário não tem acesso aos Cards.";
                    return null;
                }

                if (i == 0)
                {
                    ERRO = _driver!.FindElement(elems[0]).Text;
                    return null;
                }

                var lista = _driver!.FindElement(elems[1]);
                var cards = lista.FindElements(By.TagName("div")).Where(c => c.GetAttribute("class")!.ToLower().Contains("cards") && 
                                                                            !c.GetAttribute("class")!.ToLower().Contains("titulo") && 
                                                                            !c.GetAttribute("class")!.ToLower().Contains("categoria")).ToList();

                var cardSelecionado = cards.Where(c => c.FindElement(By.TagName("div")).Text.ToUpper().Contains(nomeTarefa.ToUpper().Trim())).ToList();

                if (cardSelecionado.Count() == 0)
                {
                    ERRO = "Card consultado não foi encontado.";
                    return null;
                }

                if (cardSelecionado.Count() > 1)
                {
                    ERRO = "Informe um nome mais específio para identificar o Card corretamente.";
                    return null;
                }

                if (abrir)
                {
                    var btn = cardSelecionado[0].FindElements(By.TagName("li")).Single(c => c.GetAttribute("aria-label")!.ToUpper().Contains("DETALHES"));
                    PosiconaScrollElemento(btn, true);
                }

                DADOS = "Card Localizado. Detalhes em tela.";
                OK = true;
                return cardSelecionado[0];
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return null;
            }
        }
        public void MarcarFlag(string nomeTarefa, string rotuloFlag = "")
        {
            try
            {
                var card = ConsultarCard(nomeTarefa);

                if (card == null)
                    return;

                DADOS = "";
                OK = false;

                PosiconaScrollElemento(card.FindElements(By.TagName("li")).Single(c => c.GetAttribute("aria-label")!.ToUpper().Contains("FLAG")), true);

                AguardaElementoExistir(By.TagName("table"));
                Aguarde(200);
                var lista = _driver!.FindElement(By.TagName("table")).FindElement(By.TagName("tbody")).FindElements(By.TagName("tr"));

                if (string.IsNullOrWhiteSpace(rotuloFlag))
                {
                    lista[0].FindElements(By.TagName("td"))[1].FindElement(By.TagName("svg")).Click();
                    Aguarde(200);
                    _driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();
                    DADOS = "Flag foi removida da Tarefa.";
                    OK = true;
                    return;
                }
                else
                {
                    var flagSelecionada = lista.Where(c => c.FindElements(By.TagName("td"))[0].Text.ToUpper().Contains(rotuloFlag.ToUpper().Trim())).ToList();

                    if (flagSelecionada.Count() == 0)
                    {
                        ERRO = "Flag não foi encontada.";
                        return;
                    }

                    if (flagSelecionada.Count() > 1)
                    {
                        ERRO = "Informe um rótulo mais específio para identificar a Flag corretamente.";
                        return;
                    }

                    string rotulo = flagSelecionada[0].FindElements(By.TagName("td"))[0].Text.Trim();
                    PosiconaScrollElemento(flagSelecionada[0].FindElements(By.TagName("td"))[1].FindElement(By.TagName("svg")), true);
                    Aguarde(200);
                    _driver.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();

                    DADOS = $"Tarefa marcada com a Flag '{rotulo}'.";
                    OK = true;
                    return;
                }
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void AssumirTramite(string nomeTarefa, string usuario = "", string senha = "")
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(senha))
                {
                    if (EstaLogadoOUAbreLogin())
                    {
                        MudarUsuario(usuario, senha); 
                    }
                    else
                    {
                        EfetuarLogin(usuario, senha);
                    }

                    if (!OK)
                        return;

                    usuario = " '"+usuario+"'";
                }

                var card = ConsultarCard(nomeTarefa);

                if (card == null)
                    return;

                DADOS = "";
                OK = false;

                var btn = card.FindElements(By.TagName("li")).Where(c => c.GetAttribute("aria-label")!.ToUpper().Contains("ASSUMIR TRÂMITE")).ToList();

                if (btn.Count() != 1)
                {
                    ERRO = "Opção para Assumir o Trâmite não está disponível para este Card.";
                    return;
                }

                PosiconaScrollElemento(btn[0], true);
                Aguarde(200);
                _driver!.FindElements(By.TagName("button")).Single(c => c.Displayed && c.Text.ToUpper().Trim() == "SIM").Click();

                DADOS = $"Trâmite atribuído ao usuário{usuario}.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void ComecarExecucaoTramite(string nomeTarefa, string usuario = "", string senha = "")
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(senha))
                {
                    if (EstaLogadoOUAbreLogin())
                    {
                        MudarUsuario(usuario, senha);
                    }
                    else
                    {
                        EfetuarLogin(usuario, senha);
                    }

                    if (!OK)
                        return;

                    usuario = " '"+usuario+"'";
                }

                var card = ConsultarCard(nomeTarefa);

                if (card == null)
                    return;

                DADOS = "";
                OK = false;

                var btn = card.FindElements(By.TagName("li")).Where(c => c.GetAttribute("aria-label")!.ToUpper().Contains("COMEÇAR TRÂMITE")).ToList();

                if (btn.Count() != 1)
                {
                    ERRO = "Opção para Iniciar a execução do Trâmite não está disponível para este Card.";
                    return;
                }

                PosiconaScrollElemento(btn[0], true);
                Aguarde(200);
                _driver!.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();

                DADOS = $"Iniciada a execução do Trâmite pelo usuário{usuario}.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void TerminarExecucaoTramite(string nomeTarefa, string nota, string usuario = "", string senha = "")
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(senha))
                {
                    if (EstaLogadoOUAbreLogin())
                    {
                        MudarUsuario(usuario, senha);
                    }
                    else
                    {
                        EfetuarLogin(usuario, senha);
                    }

                    if (!OK)
                        return;

                    usuario = " '"+usuario+"'";
                }

                var card = ConsultarCard(nomeTarefa);

                if (card == null)
                    return;

                DADOS = "";
                OK = false;

                var btn = card.FindElements(By.TagName("li")).Where(c => c.GetAttribute("aria-label")!.ToUpper().Contains("TERMINAR EXECUÇÃO DO TRÂMITE")).ToList();

                if (btn.Count() != 1)
                {
                    ERRO = "Opção para Terminar a execução do Trâmite não está disponível para este Card.";
                    return;
                }

                PosiconaScrollElemento(btn[0], true);

                AguardaElementoExistir(By.Id("nota"));
                Digitar(_driver!.FindElement(By.Id("nota")), nota);
                
                _driver!.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "OK").Click();
                Aguarde(200);
                _driver!.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();

                DADOS = $"Finalizada a execução do Trâmite pelo usuário{usuario}.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }
        public void RevisaoTramite(string nomeTarefa, bool revisadoOk, string nota, string usuario = "", string senha = "")
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(senha))
                {
                    if (EstaLogadoOUAbreLogin())
                    {
                        MudarUsuario(usuario, senha);
                    }
                    else
                    {
                        EfetuarLogin(usuario, senha);
                    }

                    if (!OK)
                        return;

                    usuario = " '"+usuario+"'";
                }

                var card = ConsultarCard(nomeTarefa);

                if (card == null)
                    return;

                DADOS = "";
                OK = false;

                string revisao = revisadoOk ? "APROVAR TRÂMITE" : "REPROVAR TRÂMITE";

                var btn = card.FindElements(By.TagName("li")).Where(c => c.GetAttribute("aria-label")!.ToUpper().Contains(revisao)).ToList();

                if (btn.Count() != 1)
                {
                    ERRO = "Opção para Revisar Trâmite não está disponível para este Card.";
                    return;
                }

                PosiconaScrollElemento(btn[0], true);

                AguardaElementoExistir(By.Id("nota"));
                Digitar(_driver!.FindElement(By.Id("nota")), nota);

                _driver!.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "OK").Click();
                Aguarde(200);
                _driver!.FindElements(By.TagName("button")).Single(c => c.Text.ToUpper().Trim() == "SIM").Click();

                DADOS = $"Trâmite revisado {(revisadoOk ? "'OK'" : "'com Falha'")} pelo usuário{usuario}.";
                OK = true;
            }
            catch (Exception ex)
            {
                ERRO = ex.Message;
                return;
            }
        }

        #endregion

        #endregion
    }
}