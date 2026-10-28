using System;
using System.Configuration;
using System.Data;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV2
{
    /// <summary>
    /// Pantalla principal del reproductor: lista las canciones de TBL_SONG,
    /// permite reproducirlas y muestra el usuario autenticado.
    /// </summary>
    public partial class Song : System.Web.UI.Page, ICallbackEventHandler
    {
        private string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

        // Resultado que se devuelve al cliente cuando termina el callback.
        private string strCallbackResult = string.Empty;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                LoadCurrentUser();
                LoadSongs();
                RegisterPlayCallback();
            }
        }

        /// <summary>
        /// Muestra el nombre (y la foto) del usuario que inicio sesion.
        /// </summary>
        private void LoadCurrentUser()
        {
            string userName = string.Empty;
            string photo = string.Empty;

            Entities.User user = Session[Util.Constants.USER] as Entities.User;
            if (user != null)
            {
                userName = user.UserName;
                photo = user.Photo;
            }

            if (string.IsNullOrEmpty(userName) && Context.User != null && Context.User.Identity != null)
            {
                userName = Context.User.Identity.Name;
            }

            if (string.IsNullOrEmpty(userName))
            {
                userName = "invitado";
            }

            if (string.IsNullOrEmpty(photo))
            {
                photo = "~/UserImages/default.jpg";
            }

            string photoUrl = ResolveMediaUrl(photo);
            lblUserName.Text = Server.HtmlEncode(userName);
            lblSidebarUser.Text = Server.HtmlEncode(userName);
            imgUserPhoto.Src = photoUrl;
            imgSidebarPhoto.Src = photoUrl;
        }

        /// <summary>
        /// Carga la lista de canciones desde la capa BLL y la enlaza al Repeater.
        /// </summary>
        private void LoadSongs()
        {
            DataTable dtSongs = new BLL.Song(strConnString).Read();

            if (dtSongs != null && dtSongs.Rows.Count > 0)
            {
                rptSongs.DataSource = dtSongs;
                rptSongs.DataBind();
                pnlEmpty.Visible = false;
                lblSummary.Text = dtSongs.Rows.Count == 1 ? "1 cancion" : dtSongs.Rows.Count + " canciones";
            }
            else
            {
                pnlEmpty.Visible = true;
                lblSummary.Text = "0 canciones";
            }
        }

        /// <summary>
        /// Registra la funcion JavaScript que avisa al servidor cuando una cancion
        /// empieza a reproducirse (para incrementar el contador de reproducciones).
        /// </summary>
        private void RegisterPlayCallback()
        {
            string callbackReference = Page.ClientScript.GetCallbackEventReference(
                this, "arg", "songPlayedCallback", "ctx", true);

            string script = "function notifyPlay(id, row) { var arg = id, ctx = row; " + callbackReference + "; }";
            Page.ClientScript.RegisterClientScriptBlock(this.GetType(), "notifyPlay", script, true);
        }

        #region ICallbackEventHandler

        public string GetCallbackResult()
        {
            return strCallbackResult;
        }

        public void RaiseCallbackEvent(string eventArgument)
        {
            strCallbackResult = string.Empty;

            int idSong;
            if (!int.TryParse(eventArgument, out idSong) || idSong <= 0)
            {
                return;
            }

            try
            {
                int plays = new BLL.Song(strConnString).IncrementPlays(idSong);
                strCallbackResult = plays.ToString();
            }
            catch (Exception)
            {
                // Si falla el incremento no se interrumpe la reproduccion.
                strCallbackResult = string.Empty;
            }
        }

        #endregion

        #region Helpers usados desde el marcado (Song.aspx)

        /// <summary>
        /// Convierte una ruta del tipo "~/songFiles/x.mp3" en una URL usable por el navegador.
        /// </summary>
        protected string ResolveMediaUrl(object path)
        {
            string value = Convert.ToString(path);
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            try
            {
                return ResolveUrl(value);
            }
            catch (Exception)
            {
                return value;
            }
        }

        /// <summary>
        /// Primera letra del nombre de la cancion (se usa como portada).
        /// </summary>
        protected string InitialOf(object value)
        {
            string name = Convert.ToString(value);
            if (string.IsNullOrEmpty(name) || name.Trim().Length == 0)
            {
                return "#";
            }
            return name.Trim().Substring(0, 1).ToUpperInvariant();
        }

        /// <summary>
        /// Muestra el contador de reproducciones, tratando los valores NULL como cero.
        /// </summary>
        protected string PlaysText(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return "0";
            }
            return Convert.ToInt32(value).ToString();
        }

        #endregion
    }
}
