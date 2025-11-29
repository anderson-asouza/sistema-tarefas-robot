using OpenQA.Selenium.BiDi.Network;
using robots;
using System.CodeDom.Compiler;
using System.Diagnostics;
using System.Runtime.Intrinsics.X86;
using System.Windows.Forms;
using static System.Runtime.InteropServices.JavaScript.JSType;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace SistemaTarefas
{
    public partial class frmPrincipal : Form
    {
        static public Robo? _robo = null;
        static DateTime _horaInicial = DateTime.Now;
        static bool _processamentoSequencial = false;

        public frmPrincipal()
        {
            InitializeComponent();

            #region ToolTip

            const string CONSULTA = "Buscar registros; necessário para atualizar ou excluir";
            const string CADASTRO = "Usado para Cadastro e Alteração";
            const string VINCULO = "Relaciona este registro à entidade pai";
            const string SENHAOPCIONAL = "Sem senha irá usar o usuário já logado.";

            #region Consulta
            toolTipFrmPrincipal.SetToolTip(txtNomeModeloTarefaCONSULTA, CONSULTA);
            toolTipFrmPrincipal.SetToolTip(txtNomeTramiteModeloTarefaCONSULTA, CONSULTA);
            toolTipFrmPrincipal.SetToolTip(txtNomeUsuarioCONSULTA, CONSULTA);
            toolTipFrmPrincipal.SetToolTip(txtNomeFlagCONSULTA, CONSULTA);
            toolTipFrmPrincipal.SetToolTip(txtNomeTramiteCONSULTA, CONSULTA);
            #endregion

            #region Cadastro
            toolTipFrmPrincipal.SetToolTip(txtNomeModeloTarefaCADASTRO, CADASTRO);
            toolTipFrmPrincipal.SetToolTip(txtNomeTramiteModeloTarefaCADASTRO, CADASTRO);
            toolTipFrmPrincipal.SetToolTip(txtNomeUsuarioCADASTRO, CADASTRO);
            toolTipFrmPrincipal.SetToolTip(txtNomeFlagCADASTRO, CADASTRO);
            #endregion

            #region Vínculo
            toolTipFrmPrincipal.SetToolTip(txtNomeModeloTarefaVINCULO, VINCULO);
            toolTipFrmPrincipal.SetToolTip(txtNomeTarefaVINCULO, VINCULO);
            toolTipFrmPrincipal.SetToolTip(txtTramiteNomeTarefaVINCULO, VINCULO);
            #endregion

            #region Senha Opcional
            toolTipFrmPrincipal.SetToolTip(txtSenhaUsuarioTramitador, SENHAOPCIONAL);
            toolTipFrmPrincipal.SetToolTip(txtSenhaUsuarioRevisor, SENHAOPCIONAL);
            #endregion

            #endregion
        }
        private void Tempo(bool inicializa = false)
        {
            if (inicializa)
            {
                _horaInicial = DateTime.Now;
                statusLabel.Text = _horaInicial.ToString("HH:mm:ss.fff");
                Application.DoEvents();
            }
            else
            {
                DateTime agora = DateTime.Now;
                TimeSpan decorrido = agora - _horaInicial;

                statusLabel.Text = $"{_horaInicial:HH:mm:ss.fff}  -  {agora:HH:mm:ss.fff}  -  ({decorrido:hh\\:mm\\:ss\\.fff})";
            }
        }
        private void AjustaLblDescritivos()
        {
            Control[] controles = { lblIntroducao, lblRetornoInicio, lblRetornoLogin, lblRetornoUsuario, lblRetornoModeloTarefa, lblRetornoTramiteModeloTarefa, lblRetornoFlag, lblRetornoTarefa, lblRetornoCard };

            foreach (var c in controles)
                c.Width = tabControles.Width - (c.Left +10);
        }
        private void frmPrincipal_SizeChanged(object sender, EventArgs e)
        {
            tabControles.Width = this.Width - ((tabControles.Left * 2) +15);
            tabControles.Height = this.Height - ((tabControles.Top * 2) +50);

            AjustaLblDescritivos();
        }
        private void frmPrincipal_Load(object sender, EventArgs e)
        {
            frmPrincipal_SizeChanged(sender, e);
        }
        private void Retorno(Label lblStatus, Label lblRetorno, bool inicializa = false)
        {
            if (inicializa)
            {
                lblStatus.Text = "- - -";
                lblStatus.ForeColor = Color.Black;
                lblRetorno.Text = "AGUARDE ...";
                Application.DoEvents();
            }
            else
            {
                if (_robo != null)
                {
                    lblStatus.Text = (_robo.OK && _robo.ERRO != "") ? "Alerta." : (_robo.OK) ? "Sucesso !!!" : "Falha!";
                    lblStatus.ForeColor = (_robo.OK && _robo.ERRO != "") ? Color.Orange : (_robo.OK) ? Color.Blue : Color.Red;

                    if (_robo.DADOS != "" && _robo.ERRO != "")
                    {
                        lblRetorno.Text = _robo.DADOS.Trim() + Environment.NewLine +
                                          "- - - ERRO - - -" + Environment.NewLine +
                                          _robo.ERRO.Trim();
                    }
                    else
                    {
                        lblRetorno.Text = _robo.DADOS + _robo.ERRO;
                    }
                }
                else
                {
                    lblStatus.Text = "NULO";
                    lblStatus.ForeColor = Color.Black;
                    lblRetorno.Text = "SEM RESPOSTA";
                }
            }
        }
        private bool ComponentesVisuais(bool acao = false, Button? DestacarBotao = null)
        {
            try
            {
                if (!acao && _robo == null)
                {
                    MessageBox.Show("Robot não foi iniciado! Inicialize o robot na Aba 'Início'.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                if (DestacarBotao != null)
                {
                    if (!acao)
                    {
                        DestacarBotao.FlatStyle = FlatStyle.Flat;
                        DestacarBotao.BackColor = SystemColors.Control;
                        DestacarBotao.FlatAppearance.BorderSize = 3;
                        DestacarBotao.FlatAppearance.BorderColor = Color.Red;
                    }
                    else
                    {
                        DestacarBotao.FlatStyle = FlatStyle.Standard;
                        DestacarBotao.BackColor = Color.Transparent;
                        DestacarBotao.FlatAppearance.BorderSize = 1;
                        DestacarBotao.FlatAppearance.BorderColor = SystemColors.ControlDark;
                    }
                }

                Control[] controles = { btnIniciarWebRobot, btnMultiplasAutomacoes, radioChrome, radioFirefox, radioEdge,
                                    btnLogar, btnLogoff, btnAlterarSenha, btnMudarUsuario, btnAdicionarFoto, btnRemoverFoto,
                                    btnCadastrarUsuario, btnConsultarUsuario, btnAtualizarUsuario, btnExcluirUsuario, btnAlterarSenhaPeloAdm, 
                                    btnCadastrarModeloTarefa, btnConsultarModeloTarefa, btnAtualizarModeloTarefa, btnExcluirModeloTarefa,
                                    btnCadastrarTramiteModeloTarefa, btnConsultarTramiteModeloTarefa, btnAtualizarTramiteModeloTarefa, btnExcluirTramiteModeloTarefa,
                                    btnCadastrarFlag, btnConsultarFlag, btnAtualizarFlag, btnExcluirFlag,
                                    btnCadastrarTarefa, btnConsultarTarefa, btnAtualizarTarefa, btnAtivarTarefa, btnDesativarTarefa, btnExcluirTarefa,
                                    btnIncluirTramite, btnVoltarTramite, btnConsultarTramite,
                                    btnConsultarCard, btnMarcarFlag, btnRemoverFlag, btnAssumirTramite, btnComecarTramite, btnFinalizarTramite, btnRevisarTramite, radioRevisadoOk, radioRevisadoFalha};

                foreach (var c in controles)
                {
                    c.Enabled = acao;
                }

                Application.DoEvents();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }
        private void btnIniciarWebRobot_Click(object sender, EventArgs e)
        {
            try
            {
                btnIniciarWebRobot.Enabled = false;
                btnMultiplasAutomacoes.Enabled = false;

                if (_robo != null)
                {
                    _robo.FecharRobo();
                    _robo = null;
                }

                Retorno(lblStatusInicio, lblRetornoInicio, true);

                int qualwebdriver = (radioChrome.Checked) ? 1 : (radioFirefox.Checked) ? 2 : 3;
                _robo = new Robo(qualwebdriver);

                ComponentesVisuais();

                Tempo(true);
                _robo.AbrirWebSite(txtUrl.Text.Trim());
                Tempo();

                Retorno(lblStatusInicio, lblRetornoInicio);
            }
            finally
            {
                ComponentesVisuais(true);
            }
        }
        private async void btnLogar_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnLogar))
                    return;

                Retorno(lblStatusLogin, lblRetornoLogin, true);

                Tempo(true);
                _robo!.EfetuarLogin(txtLogin.Text, txtSenha.Text, txtNovaSenha.Text, txtConfirmacaoNovaSenha.Text);
                Tempo();

                Retorno(lblStatusLogin, lblRetornoLogin);
            }
            finally
            {
                ComponentesVisuais(true, btnLogar);
            }
        }
        private void btnLogoff_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnLogoff))
                    return;

                Retorno(lblStatusLogin, lblRetornoLogin, true);

                Tempo(true);
                _robo!.EfetuarLogoff();
                Tempo();

                Retorno(lblStatusLogin, lblRetornoLogin);
            }
            finally
            {
                ComponentesVisuais(true, btnLogoff);
            }
        }
        private void btnMudarUsuario_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnMudarUsuario))
                    return;

                Retorno(lblStatusLogin, lblRetornoLogin, true);

                Tempo(true);
                _robo!.MudarUsuario(txtLogin.Text, txtSenha.Text);
                Tempo();

                Retorno(lblStatusLogin, lblRetornoLogin);
            }
            finally
            {
                ComponentesVisuais(true, btnMudarUsuario);
            }
        }
        private void btnAdicionarFoto_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAdicionarFoto))
                    return;

                Retorno(lblStatusLogin, lblRetornoLogin, true);

                Tempo(true);
                _robo!.AdicionarFoto(txtFoto.Text);
                Tempo();

                Retorno(lblStatusLogin, lblRetornoLogin);
            }
            finally
            {
                ComponentesVisuais(true, btnAdicionarFoto);
            }
        }
        private void btnRemoverFoto_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnRemoverFoto))
                    return;

                Retorno(lblStatusLogin, lblRetornoLogin, true);

                Tempo(true);
                _robo!.RemoverFoto();
                Tempo();

                Retorno(lblStatusLogin, lblRetornoLogin);
            }
            finally
            {
                ComponentesVisuais(true, btnRemoverFoto);
            }
        }
        private void btnAlterarSenha_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAlterarSenha))
                    return;

                Retorno(lblStatusLogin, lblRetornoLogin, true);

                Tempo(true);
                _robo!.AlterarSenha(txtLogin.Text, txtSenha.Text, txtNovaSenha.Text, txtConfirmacaoNovaSenha.Text);
                Tempo();

                Retorno(lblStatusLogin, lblRetornoLogin);
            }
            finally
            {
                ComponentesVisuais(true, btnAlterarSenha);
            }
        }
        private void btnCadastrarModeloTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnCadastrarModeloTarefa))
                    return;

                Retorno(lblStatusModeloTarefa, lblRetornoModeloTarefa, true);

                Tempo(true);
                _robo!.CadastroModeloTarefa(true, "", txtNomeModeloTarefaCADASTRO.Text, txtDescricaoModeloTarefa.Text);
                Tempo();

                Retorno(lblStatusModeloTarefa, lblRetornoModeloTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnCadastrarModeloTarefa);
            }
        }
        private void btnConsultarModeloTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnConsultarModeloTarefa))
                    return;

                Retorno(lblStatusModeloTarefa, lblRetornoModeloTarefa, true);

                Tempo(true);
                _robo!.ConsultarModeloTarefa(txtNomeModeloTarefaCONSULTA.Text);
                Tempo();

                Retorno(lblStatusModeloTarefa, lblRetornoModeloTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnConsultarModeloTarefa);
            }
        }
        private void btnAtualizarModeloTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAtualizarModeloTarefa))
                    return;

                Retorno(lblStatusModeloTarefa, lblRetornoModeloTarefa, true);

                Tempo(true);
                _robo!.CadastroModeloTarefa(false, txtNomeModeloTarefaCONSULTA.Text, txtNomeModeloTarefaCADASTRO.Text, txtDescricaoModeloTarefa.Text);
                Tempo();

                Retorno(lblStatusModeloTarefa, lblRetornoModeloTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnAtualizarModeloTarefa);
            }
        }
        private void btnExcluirModeloTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("Deseja realmente excluir este cadastro?",
                                                      "Confirmação",
                                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                    return;

                if (!ComponentesVisuais(false, btnExcluirModeloTarefa))
                    return;

                Retorno(lblStatusModeloTarefa, lblRetornoModeloTarefa, true);

                Tempo(true);
                _robo!.ExcluirModeloTarefa(txtNomeModeloTarefaCONSULTA.Text);
                Tempo();

                Retorno(lblStatusModeloTarefa, lblRetornoModeloTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnExcluirModeloTarefa);
            }
        }
        private void btnCadastrarTramiteModeloTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnCadastrarTramiteModeloTarefa))
                    return;

                Retorno(lblStatusTramiteModeloTarefa, lblRetornoTramiteModeloTarefa, true);

                Tempo(true);
                _robo!.CadastroTramiteModeloTarefa(true, txtNomeModeloTarefaVINCULO.Text, "",
                    txtNomeTramiteModeloTarefaCADASTRO.Text, txtDescricaoTramiteModeloTarefa.Text, txtDuracao.Value.ToString(), txtUsuarioRevisor.Text, txtUsuarioTramitador.Text);
                Tempo();

                Retorno(lblStatusTramiteModeloTarefa, lblRetornoTramiteModeloTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnCadastrarTramiteModeloTarefa);
            }
        }
        private void btnConsultarTramiteModeloTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtNomeModeloTarefaVINCULO.Text.Trim() == "" || txtNomeTramiteModeloTarefaCONSULTA.Text.Trim() == "")
                {
                    MessageBox.Show("Para Consultar precisa preencher o Vínculo e Trâmite Consulta.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!ComponentesVisuais(false, btnConsultarTramiteModeloTarefa))
                    return;

                Retorno(lblStatusTramiteModeloTarefa, lblRetornoTramiteModeloTarefa, true);

                Tempo(true);
                _robo!.ConsultarTramiteModeloTarefa(txtNomeModeloTarefaVINCULO.Text, txtNomeTramiteModeloTarefaCONSULTA.Text);
                Tempo();

                Retorno(lblStatusTramiteModeloTarefa, lblRetornoTramiteModeloTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnConsultarTramiteModeloTarefa);
            }
        }
        private void btnAtualizarTramiteModeloTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAtualizarTramiteModeloTarefa))
                    return;

                Retorno(lblStatusTramiteModeloTarefa, lblRetornoTramiteModeloTarefa, true);

                Tempo(true);
                _robo!.CadastroTramiteModeloTarefa(false, txtNomeModeloTarefaVINCULO.Text, txtNomeTramiteModeloTarefaCONSULTA.Text,
                    txtNomeTramiteModeloTarefaCADASTRO.Text, txtDescricaoTramiteModeloTarefa.Text, txtDuracao.Value.ToString(), txtUsuarioRevisor.Text, txtUsuarioTramitador.Text);
                Tempo();

                Retorno(lblStatusTramiteModeloTarefa, lblRetornoTramiteModeloTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnAtualizarTramiteModeloTarefa);
            }
        }
        private void btnExcluirTramiteModeloTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("Deseja realmente excluir este cadastro?",
                                                      "Confirmação",
                                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                    return;

                if (!ComponentesVisuais(false, btnExcluirTramiteModeloTarefa))
                    return;

                Retorno(lblStatusTramiteModeloTarefa, lblRetornoTramiteModeloTarefa, true);

                Tempo(true);
                _robo!.ExcluirTramiteModeloTarefa(txtNomeModeloTarefaVINCULO.Text, txtNomeTramiteModeloTarefaCONSULTA.Text);

                Tempo();

                Retorno(lblStatusTramiteModeloTarefa, lblRetornoTramiteModeloTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnExcluirTramiteModeloTarefa);
            }
        }
        private void btnCadastrarUsuario_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnExcluirTramiteModeloTarefa))
                    return;

                Retorno(lblStatusUsuario, lblRetornoUsuario, true);

                Tempo(true);
                _robo!.CadastroUsuario(true, "", txtNomeUsuarioCADASTRO.Text, txtLoginUsuario.Text, txtSenhaInicialUsuario.Text, txtNivelUsuario.Text, txtEmailUsuario.Text, txtMatriculaUsuario.Text);
                Tempo();

                Retorno(lblStatusUsuario, lblRetornoUsuario);
            }
            finally
            {
                ComponentesVisuais(true, btnExcluirTramiteModeloTarefa);
            }
        }
        private void btnConsultarUsuario_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnConsultarUsuario))
                    return;

                Retorno(lblStatusUsuario, lblRetornoUsuario, true);

                Tempo(true);
                _robo!.ConsultarUsuario(txtNomeUsuarioCONSULTA.Text);
                Tempo();

                Retorno(lblStatusUsuario, lblRetornoUsuario);
            }
            finally
            {
                ComponentesVisuais(true, btnConsultarUsuario);
            }
        }
        private void btnAtualizarUsuario_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAtualizarUsuario))
                    return;

                Retorno(lblStatusUsuario, lblRetornoUsuario, true);

                Tempo(true);
                _robo!.CadastroUsuario(false, txtNomeUsuarioCONSULTA.Text, txtNomeUsuarioCADASTRO.Text, "", "", txtNivelUsuario.Text, txtEmailUsuario.Text, txtMatriculaUsuario.Text);
                Tempo();

                Retorno(lblStatusUsuario, lblRetornoUsuario);
            }
            finally
            {
                ComponentesVisuais(true, btnAtualizarUsuario);
            }
        }
        private void btnExcluirUsuario_Click(object sender, EventArgs e)
        {
            try
            {
                if (!_processamentoSequencial)
                {
                    DialogResult result = MessageBox.Show("Deseja realmente excluir este cadastro?",
                                          "Confirmação",
                                          MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                        return;
                }

                if (!ComponentesVisuais(false, btnExcluirUsuario))
                    return;

                Retorno(lblStatusUsuario, lblRetornoUsuario, true);

                Tempo(true);
                _robo!.ExcluirUsuario(txtNomeUsuarioCONSULTA.Text);
                Tempo();

                Retorno(lblStatusUsuario, lblRetornoUsuario);
            }
            finally
            {
                ComponentesVisuais(true, btnExcluirUsuario);
            }
        }
        private void btnAlterarSenhaPeloAdm_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAlterarSenhaPeloAdm))
                    return;

                Retorno(lblStatusLogin, lblRetornoLogin, true);

                Tempo(true);
                _robo!.AlterarSenhaPeloAdm(txtLoginAlterarSenhaPeloAdm.Text, txtNovaSenhaAlterarSenhaPeloAdm.Text, txtConfirmacaoAlterarSenhaPeloAdm.Text);
                Tempo();

                Retorno(lblStatusLogin, lblRetornoLogin);
            }
            finally
            {
                ComponentesVisuais(true, btnAlterarSenhaPeloAdm);
            }
        }
        private void btnCadastrarFlag_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnCadastrarFlag))
                    return;

                Retorno(lblStatusFlag, lblRetornoFlag, true);

                Tempo(true);
                _robo!.CadastroFlag(true, "", txtNomeFlagCADASTRO.Text, txtCorFlag.Text);
                Tempo();

                Retorno(lblStatusFlag, lblRetornoFlag);
            }
            finally
            {
                ComponentesVisuais(true, btnCadastrarFlag);
            }
        }
        private void btnConsultarFlag_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnConsultarFlag))
                    return;

                Retorno(lblStatusFlag, lblRetornoFlag, true);

                Tempo(true);
                _robo!.ConsultarFlag(txtNomeFlagCONSULTA.Text);
                Tempo();

                Retorno(lblStatusFlag, lblRetornoFlag);
            }
            finally
            {
                ComponentesVisuais(true, btnConsultarFlag);
            }
        }
        private void btnAtualizarFlag_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAtualizarFlag))
                    return;

                Retorno(lblStatusFlag, lblRetornoFlag, true);

                Tempo(true);
                _robo!.CadastroFlag(false, txtNomeFlagCONSULTA.Text, txtNomeFlagCADASTRO.Text, txtCorFlag.Text);
                Tempo();

                Retorno(lblStatusFlag, lblRetornoFlag);
            }
            finally
            {
                ComponentesVisuais(true, btnAtualizarFlag);
            }
        }
        private void btnExcluirFlag_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("Deseja realmente excluir este cadastro?",
                                      "Confirmação",
                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                    return;

                if (!ComponentesVisuais(false, btnExcluirFlag))
                    return;

                Retorno(lblStatusFlag, lblRetornoFlag, true);

                Tempo(true);
                _robo!.ExcluirFlag(txtNomeFlagCONSULTA.Text);
                Tempo();

                Retorno(lblStatusFlag, lblRetornoFlag);
            }
            finally
            {
                ComponentesVisuais(true, btnExcluirFlag);
            }
        }
        private void btnCadastrarTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnCadastrarTarefa))
                    return;

                Retorno(lblStatusTarefa, lblRetornoTarefa, true);

                Tempo(true);
                _robo!.CadastroTarefa(true, txtNomeTarefaVINCULO.Text, "", txtNomeTarefaCADASTRO.Text, txtDescricaoTarefa.Text);
                Tempo();

                Retorno(lblStatusTarefa, lblRetornoTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnCadastrarTarefa);
            }
        }

        private void btnConsultarTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnConsultarTarefa))
                    return;

                Retorno(lblStatusTarefa, lblRetornoTarefa, true);

                Tempo(true);
                _robo!.ConsultarTarefa(txtNomeTarefaCONSULTA.Text);
                Tempo();

                Retorno(lblStatusTarefa, lblRetornoTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnConsultarTarefa);
            }
        }
        private void btnAtualizarTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAtualizarTarefa))
                    return;

                Retorno(lblStatusTarefa, lblRetornoTarefa, true);

                Tempo(true);
                _robo!.CadastroTarefa(false, "", txtNomeTarefaCONSULTA.Text, txtNomeTarefaCADASTRO.Text, txtDescricaoTarefa.Text);
                Tempo();

                Retorno(lblStatusTarefa, lblRetornoTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnAtualizarTarefa);
            }
        }
        private void btnAtivarTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAtivarTarefa))
                    return;

                Retorno(lblStatusTarefa, lblRetornoTarefa, true);

                Tempo(true);
                _robo!.AtivarDesativarTarefa(txtNomeTarefaCONSULTA.Text, true);
                Tempo();

                Retorno(lblStatusTarefa, lblRetornoTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnAtivarTarefa);
            }
        }
        private void btnDesativarTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnDesativarTarefa))
                    return;

                Retorno(lblStatusTarefa, lblRetornoTarefa, true);

                Tempo(true);
                _robo!.AtivarDesativarTarefa(txtNomeTarefaCONSULTA.Text, false);
                Tempo();

                Retorno(lblStatusTarefa, lblRetornoTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnDesativarTarefa);
            }
        }
        private void btnExcluirTarefa_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = MessageBox.Show("Deseja realmente excluir este cadastro?",
                                                      "Confirmação",
                                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                    return;

                if (!ComponentesVisuais(false, btnExcluirTarefa))
                    return;

                Retorno(lblStatusTarefa, lblRetornoTarefa, true);

                Tempo(true);
                _robo!.ExcluirTarefa(txtNomeTarefaCONSULTA.Text);
                Tempo();

                Retorno(lblStatusTarefa, lblRetornoTarefa);
            }
            finally
            {
                ComponentesVisuais(true, btnExcluirTarefa);
            }
        }
        private void btnIncluirTramite_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnIncluirTramite))
                    return;

                Retorno(lblStatusTramite, lblRetornoTramite, true);

                Tempo(true);
                _robo!.IncluirTramite(txtTramiteNomeTarefaVINCULO.Text);
                Tempo();

                Retorno(lblStatusTramite, lblRetornoTramite);
            }
            finally
            {
                ComponentesVisuais(true, btnIncluirTramite);
            }
        }
        private void btnVoltarTramite_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnVoltarTramite))
                    return;

                Retorno(lblStatusTramite, lblRetornoTramite, true);

                Tempo(true);
                _robo!.VoltarTramite(txtTramiteNomeTarefaVINCULO.Text, txtNomeTramiteCONSULTA.Text);
                Tempo();

                Retorno(lblStatusTramite, lblRetornoTramite);
            }
            finally
            {
                ComponentesVisuais(true, btnVoltarTramite);
            }
        }
        private void btnConsultarTramite_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnConsultarTramite))
                    return;

                Retorno(lblStatusTramite, lblRetornoTramite, true);

                Tempo(true);
                _robo!.ConsultarTramite(txtTramiteNomeTarefaVINCULO.Text, txtNomeTramiteCONSULTA.Text);
                Tempo();

                Retorno(lblStatusTramite, lblRetornoTramite);
            }
            finally
            {
                ComponentesVisuais(true, btnConsultarTramite);
            }
        }

        private void btnConsultarCard_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnConsultarCard))
                    return;

                Retorno(lblStatusCard, lblRetornoCard, true);

                Tempo(true);
                _robo!.ConsultarCard(txtCardNomeTarefa.Text, true);
                Tempo();

                Retorno(lblStatusCard, lblRetornoCard);
            }
            finally
            {
                ComponentesVisuais(true, btnConsultarCard);
            }
        }
        private void btnMarcarFlag_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnMarcarFlag))
                    return;

                Retorno(lblStatusCard, lblRetornoCard, true);

                Tempo(true);
                _robo!.MarcarFlag(txtCardNomeTarefa.Text, txtCardRotuloFlag.Text);
                Tempo();

                Retorno(lblStatusCard, lblRetornoCard);
            }
            finally
            {
                ComponentesVisuais(true, btnMarcarFlag);
            }
        }

        private void btnRemoverFlag_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnMarcarFlag))
                    return;

                Retorno(lblStatusCard, lblRetornoCard, true);

                Tempo(true);
                _robo!.MarcarFlag(txtCardNomeTarefa.Text);
                Tempo();

                Retorno(lblStatusCard, lblRetornoCard);
            }
            finally
            {
                ComponentesVisuais(true, btnMarcarFlag);
            }
        }

        private void btnAssumirTramite_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnAssumirTramite))
                    return;

                Retorno(lblStatusCard, lblRetornoCard, true);

                Tempo(true);
                _robo!.AssumirTramite(txtCardNomeTarefa.Text, txtCardTramitador.Text, txtSenhaUsuarioTramitador.Text);
                Tempo();

                Retorno(lblStatusCard, lblRetornoCard);
            }
            finally
            {
                ComponentesVisuais(true, btnAssumirTramite);
            }
        }

        private void btnComecarTramite_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnComecarTramite))
                    return;

                Retorno(lblStatusCard, lblRetornoCard, true);

                Tempo(true);
                _robo!.ComecarExecucaoTramite(txtCardNomeTarefa.Text, txtCardTramitador.Text, txtSenhaUsuarioTramitador.Text);
                Tempo();

                Retorno(lblStatusCard, lblRetornoCard);
            }
            finally
            {
                ComponentesVisuais(true, btnComecarTramite);
            }
        }
        private void btnFinalizarTramite_Click(object sender, EventArgs e)
        {
            try
            {
                if (!ComponentesVisuais(false, btnFinalizarTramite))
                    return;

                Retorno(lblStatusCard, lblRetornoCard, true);

                Tempo(true);
                _robo!.TerminarExecucaoTramite(txtCardNomeTarefa.Text, txtNotaExecucao.Text, txtCardTramitador.Text, txtSenhaUsuarioTramitador.Text);
                Tempo();

                Retorno(lblStatusCard, lblRetornoCard);
            }
            finally
            {
                ComponentesVisuais(true, btnFinalizarTramite);
            }
        }

        private void btnRevisarTramite_Click(object sender, EventArgs e)
        {
            try
            {
                if (!radioRevisadoOk.Checked && !radioRevisadoFalha.Checked)
                {
                    MessageBox.Show("Deve informar se a Revisão está OK ou com falha.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!ComponentesVisuais(false, btnRevisarTramite))
                    return;

                Retorno(lblStatusCard, lblRetornoCard, true);

                Tempo(true);
                _robo!.RevisaoTramite(txtCardNomeTarefa.Text, radioRevisadoOk.Checked, txtNotaRevisao.Text, txtCardRevisor.Text, txtSenhaUsuarioRevisor.Text);
                Tempo();

                Retorno(lblStatusCard, lblRetornoCard);
            }
            finally
            {
                ComponentesVisuais(true, btnRevisarTramite);
            }
        }

        private void btnMultiplasAutomacoes_Click(object sender, EventArgs e)
        {
            // Nota: A adição de foto de perfil é muito rápida, 
            // então o usuário pode não perceber a mudança. 
            // Para testar a funcionalidade e não gerar confusão,
            // a foto é adicionada, removida e restaurada em seguida.

            string tramitador = "";
            string sControle = "";

            try
            {
                DialogResult result = MessageBox.Show("Executar automações em sequência." +Environment.NewLine+Environment.NewLine+
                                                      "Algumas ações dependem do estado resultante das etapas anteriores. "+
                                                      "Por esse motivo, certas operações não serão executadas (como 'Excluir')."+Environment.NewLine+Environment.NewLine+
                                                      "Outras ações, como 'Login', aparecem mais de uma vez, pois são necessárias "+
                                                      "para continuar a sequência após operações como o 'Logoff'." +Environment.NewLine +Environment.NewLine+
                                                      "Deseja continuar ?",
                                      "Confirmação",
                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.No)
                    return;

                _processamentoSequencial = true;

                if (!ComponentesVisuais())
                    return;

                Retorno(lblStatusInicio, lblRetornoInicio, true);

                DateTime horaComeco = DateTime.Now;
                int i = 0;

                Control[] controles = { btnLogar, btnLogoff, btnLogar, btnMudarUsuario, btnAlterarSenha, btnAdicionarFoto, btnRemoverFoto, btnAdicionarFoto,
                                    btnCadastrarUsuario, btnConsultarUsuario, btnAtualizarUsuario, btnAlterarSenhaPeloAdm, btnExcluirUsuario,
                                    btnCadastrarModeloTarefa, btnConsultarModeloTarefa, btnAtualizarModeloTarefa,
                                    btnCadastrarTramiteModeloTarefa, btnConsultarTramiteModeloTarefa, btnAtualizarTramiteModeloTarefa,
                                    btnCadastrarFlag, btnConsultarFlag, btnAtualizarFlag,
                                    btnCadastrarTarefa, btnConsultarTarefa, btnAtualizarTarefa, btnDesativarTarefa, btnAtivarTarefa,
                                    btnVoltarTramite, btnIncluirTramite, btnConsultarTramite,
                                    btnConsultarCard, btnMarcarFlag, btnRemoverFlag, btnAssumirTramite, btnComecarTramite, btnFinalizarTramite, btnRevisarTramite};

                foreach (var c in controles)
                {
                    i++;

                    if (c is Button btn)
                        sControle = btn.Text;

                    switch (i)
                    {
                        case 1:
                            tabControles.SelectedIndex = 1;
                            break;
                        case 9:
                            tabControles.SelectedIndex = 2;
                            break;
                        case 14:
                            tabControles.SelectedIndex = 3;
                            break;
                        case 17:
                            tabControles.SelectedIndex = 4;
                            break;
                        case 20:
                            tabControles.SelectedIndex = 5;
                            break;
                        case 23:
                            tabControles.SelectedIndex = 6;
                            break;
                        case 28:
                            tabControles.SelectedIndex = 7;
                            break;
                        case 31:
                            tabControles.SelectedIndex = 8;
                            break;
                        case 35:
                            if (txtSenhaUsuarioTramitador.Text.Trim() != "")
                            {
                                tramitador = txtSenhaUsuarioTramitador.Text;
                                txtSenhaUsuarioTramitador.Text = "";
                            }
                            break;
                        case 37:
                            if (tramitador != "")
                            {
                                txtSenhaUsuarioTramitador.Text = tramitador;
                            }
                            break;
                    }

                    Application.DoEvents();

                    switch (i)
                    {
                        case 1:
                        case 3:
                            btnLogar_Click(sender, EventArgs.Empty);
                            break;
                        case 2:
                            btnLogoff_Click(sender, EventArgs.Empty);
                            break;
                        case 4:
                            btnMudarUsuario_Click(sender, EventArgs.Empty);
                            break;
                        case 5:
                            btnAlterarSenha_Click(sender, EventArgs.Empty);
                            break;
                        case 6:
                        case 8:
                            btnAdicionarFoto_Click(sender, EventArgs.Empty);
                            break;
                        case 7:
                            btnRemoverFoto_Click(sender, EventArgs.Empty);
                            break;
                        case 9:
                            btnCadastrarUsuario_Click(sender, EventArgs.Empty);
                            break;
                        case 10:
                            btnConsultarUsuario_Click(sender, EventArgs.Empty);
                            break;
                        case 11:
                            btnAtualizarUsuario_Click(sender, EventArgs.Empty);
                            break;
                        case 12:
                            btnAlterarSenhaPeloAdm_Click(sender, EventArgs.Empty);
                            break;
                        case 13:
                            btnExcluirUsuario_Click(sender, EventArgs.Empty);
                            break;
                        case 14:
                            btnCadastrarModeloTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 15:
                            btnConsultarModeloTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 16:
                            btnAtualizarModeloTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 17:
                            btnCadastrarTramiteModeloTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 18:
                            btnConsultarTramiteModeloTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 19:
                            btnAtualizarTramiteModeloTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 20:
                            btnCadastrarFlag_Click(sender, EventArgs.Empty);
                            break;
                        case 21:
                            btnConsultarFlag_Click(sender, EventArgs.Empty);
                            break;
                        case 22:
                            btnAtualizarFlag_Click(sender, EventArgs.Empty);
                            break;
                        case 23:
                            btnCadastrarTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 24:
                            btnConsultarTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 25:
                            btnAtualizarTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 26:
                            btnDesativarTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 27:
                            btnAtivarTarefa_Click(sender, EventArgs.Empty);
                            break;
                        case 28:
                            btnVoltarTramite_Click(sender, EventArgs.Empty);
                            break;
                        case 29:
                            btnIncluirTramite_Click(sender, EventArgs.Empty);
                            break;
                        case 30:
                            btnConsultarTramite_Click(sender, EventArgs.Empty);
                            break;
                        case 31:
                            btnConsultarCard_Click(sender, EventArgs.Empty);
                            break;
                        case 32:
                            btnMarcarFlag_Click(sender, EventArgs.Empty);
                            break;
                        case 33:
                            btnRemoverFlag_Click(sender, EventArgs.Empty);
                            break;
                        case 34:
                            btnAssumirTramite_Click(sender, EventArgs.Empty);
                            break;
                        case 35:
                            btnComecarTramite_Click(sender, EventArgs.Empty);
                            break;
                        case 36:
                            btnFinalizarTramite_Click(sender, EventArgs.Empty);
                            break;
                        case 37:
                            btnRevisarTramite_Click(sender, EventArgs.Empty);
                            break;
                    }

                    if (!_robo!.OK)
                    {
                        MessageBox.Show("Houve falha na execução de uma das automações. Verifique.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        _robo.ERRO += Environment.NewLine+"Verifique a automação: "+sControle;
                        return;
                    }
                }

                if (_robo!.OK)
                {
                    tabControles.SelectedIndex = 0;
                    _robo.DADOS = "Múltiplas Automações executadas com sucesso.";

                    _horaInicial = horaComeco;
                    Tempo();
                }                
            }
            finally
            {
                Retorno(lblStatusInicio, lblRetornoInicio);

                if (tramitador != "")
                {
                    txtSenhaUsuarioTramitador.Text = tramitador;
                }
                ComponentesVisuais(true, btnRevisarTramite);
                _processamentoSequencial = false;
            }
        }
    }
}