using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string strConnString = System.Configuration.ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
                BLL.PoliMusicV3.Song songBLL = new BLL.PoliMusicV3.Song(strConnString);
                DataTable dtSongs = songBLL.Read();
                gvSongs.DataSource = dtSongs;
                gvSongs.DataBind();
            }
        }
    }
}