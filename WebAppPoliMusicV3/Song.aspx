<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>
<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <h1>Listado de Canciones</h1>

    <asp:GridView ID="gvSongs" runat="server" CssClass="table table-striped table-hover" AutoGenerateColumns="false" GridLines="None">
        <Columns>
            <asp:BoundField DataField="ID_SONG" HeaderText="ID" />
            <asp:BoundField DataField="SONG_NAME" HeaderText="Canción" />
            <asp:BoundField DataField="SONG_PATH" HeaderText="Ruta" />
            <asp:BoundField DataField="PLAYS" HeaderText="Reproducciones" />
            <asp:TemplateField HeaderText="Reproducir">
                <ItemTemplate>
                    <button type="button" class="btn btn-sm btn-primary song-play-btn" data-state="stopped">Play</button>
                    <audio class="song-audio" preload="none" src='<%# ResolveClientUrl(Convert.ToString(Eval("SONG_PATH"))) %>'></audio>
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>
    </asp:GridView>

    <script type="text/javascript">
        (function () {
            function stopOtherPlayers(currentAudio) {
                var audios = document.querySelectorAll("#<%= gvSongs.ClientID %> .song-audio");
                for (var i = 0; i < audios.length; i++) {
                    if (audios[i] !== currentAudio) {
                        audios[i].pause();
                        var otherRow = audios[i].closest("tr");
                        if (!otherRow) {
                            continue;
                        }

                        var otherButton = otherRow.querySelector(".song-play-btn");
                        if (otherButton) {
                            otherButton.textContent = "Play";
                            otherButton.setAttribute("data-state", "stopped");
                        }
                    }
                }
            }

            document.addEventListener("click", function (event) {
                var button = event.target.closest(".song-play-btn");
                if (!button) {
                    return;
                }

                var row = button.closest("tr");
                if (!row) {
                    return;
                }

                var audio = row.querySelector(".song-audio");
                if (!audio || !audio.getAttribute("src")) {
                    return;
                }

                if (audio.paused) {
                    stopOtherPlayers(audio);
                    audio.play();
                    button.textContent = "Pause";
                    button.setAttribute("data-state", "playing");
                } else {
                    audio.pause();
                    button.textContent = "Play";
                    button.setAttribute("data-state", "stopped");
                }
            });

            document.addEventListener("ended", function (event) {
                var audio = event.target;
                if (!audio.classList.contains("song-audio")) {
                    return;
                }

                var row = audio.closest("tr");
                if (!row) {
                    return;
                }

                var button = row.querySelector(".song-play-btn");
                if (button) {
                    button.textContent = "Play";
                    button.setAttribute("data-state", "stopped");
                }
            }, true);
        })();
    </script>
</asp:Content>
