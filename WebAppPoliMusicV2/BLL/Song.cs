using System;
using System.Collections.Generic;
using System.Data;
using System.Web;

namespace WebAppPoliMusicV2.BLL
{
    /// <summary>
    /// Reglas de negocio de las canciones. Delega el acceso a datos en la capa DAL.
    /// </summary>
    public class Song
    {
        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public int Save(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Save(song);
        }

        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }

        public DataTable ReadById(int idSong)
        {
            return new DAL.Song(strConnectionString).ReadById(idSong);
        }

        public int Update(Entities.Song song)
        {
            return new DAL.Song(strConnectionString).Update(song);
        }

        public int Delete(int idSong)
        {
            return new DAL.Song(strConnectionString).Delete(idSong);
        }

        public int IncrementPlays(int idSong)
        {
            return new DAL.Song(strConnectionString).IncrementPlays(idSong);
        }

        /// <summary>
        /// Valida los datos minimos de una cancion antes de guardarla.
        /// Devuelve cadena vacia cuando la cancion es valida.
        /// </summary>
        public string Validate(Entities.Song song)
        {
            if (song == null)
            {
                return "La cancion no puede ser nula.";
            }
            if (string.IsNullOrEmpty(song.Name) || song.Name.Trim().Length == 0)
            {
                return "El nombre de la cancion es obligatorio.";
            }
            if (string.IsNullOrEmpty(song.Path) || song.Path.Trim().Length == 0)
            {
                return "La ruta del archivo es obligatoria.";
            }
            if (song.Plays < 0)
            {
                return "El numero de reproducciones no puede ser negativo.";
            }
            return string.Empty;
        }
    }

}
