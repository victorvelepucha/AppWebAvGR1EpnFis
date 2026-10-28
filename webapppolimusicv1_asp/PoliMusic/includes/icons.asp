<%
' ============================================================================
'  PoliMusic - Biblioteca de iconos SVG en linea (sin dependencias externas)
'  Todos usan currentColor para heredar el color del contenedor.
' ============================================================================

Function Ico(name, size)
    Dim s, sz, p
    s = LCase(Nz(name, "music"))
    If ToInt(size) <= 0 Then
        sz = 20
    Else
        sz = ToInt(size)
    End If
    p = "<svg class=""ico ico-" & H(s) & """ width=""" & sz & """ height=""" & sz & _
        """ viewBox=""0 0 24 24"" fill=""none"" aria-hidden=""true"" focusable=""false"">"

    Select Case s
        Case "home"
            p = p & "<path d=""M3 10.2 12 3l9 7.2"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<path d=""M5.5 9.6V20a1 1 0 0 0 1 1h3.2v-5.6h4.6V21h3.2a1 1 0 0 0 1-1V9.6"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "search"
            p = p & "<circle cx=""10.8"" cy=""10.8"" r=""6.3"" stroke=""currentColor"" stroke-width=""1.9""/>" & _
                    "<path d=""m15.6 15.6 4.4 4.4"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round""/>"
        Case "library"
            p = p & "<path d=""M4 4h2.6v16H4zM9.2 4h2.6v16H9.2z"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linejoin=""round""/>" & _
                    "<path d=""m15.4 5.1 2.5-.6 4 14.2-2.5.7z"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linejoin=""round""/>"
        Case "plus"
            p = p & "<path d=""M12 5v14M5 12h14"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round""/>"
        Case "heart"
            p = p & "<path d=""M12 20s-7.5-4.6-7.5-9.6A4.4 4.4 0 0 1 12 7.6a4.4 4.4 0 0 1 7.5 2.8C19.5 15.4 12 20 12 20Z"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linejoin=""round""/>"
        Case "heart-fill"
            p = p & "<path d=""M12 20.3s-8-4.9-8-10.2A4.7 4.7 0 0 1 12 7.2a4.7 4.7 0 0 1 8 2.9c0 5.3-8 10.2-8 10.2Z"" fill=""currentColor""/>"
        Case "play"
            p = p & "<path d=""M8 5.4v13.2a.8.8 0 0 0 1.2.7l10.4-6.6a.8.8 0 0 0 0-1.4L9.2 4.7A.8.8 0 0 0 8 5.4Z"" fill=""currentColor""/>"
        Case "pause"
            p = p & "<path d=""M7.5 5h3.2v14H7.5zM13.3 5h3.2v14h-3.2z"" fill=""currentColor""/>"
        Case "next"
            p = p & "<path d=""M6 5.7v12.6a.7.7 0 0 0 1.1.6l9-6.3a.7.7 0 0 0 0-1.2l-9-6.3A.7.7 0 0 0 6 5.7Z"" fill=""currentColor""/>" & _
                    "<path d=""M17.4 5h2.1v14h-2.1z"" fill=""currentColor""/>"
        Case "prev"
            p = p & "<path d=""M18 5.7v12.6a.7.7 0 0 1-1.1.6l-9-6.3a.7.7 0 0 1 0-1.2l9-6.3a.7.7 0 0 1 1.1.6Z"" fill=""currentColor""/>" & _
                    "<path d=""M4.5 5h2.1v14H4.5z"" fill=""currentColor""/>"
        Case "shuffle"
            p = p & "<path d=""M17 4.5 20.5 8 17 11.5M17 12.5 20.5 16 17 19.5"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<path d=""M3.5 8h3.2c1.6 0 2.6 1 3.6 2.4m0 3.2c-1 1.4-2 2.4-3.6 2.4H3.5"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>" & _
                    "<path d=""M13.6 10.4c1-1.4 2-2.4 3.4-2.4h3.5"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>"
        Case "repeat"
            p = p & "<path d=""M6.5 8.5h9.8l-2.4-2.6M17.5 15.5H7.7l2.4 2.6"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<path d=""M20 11.2V10a2 2 0 0 0-2-2h-.6M4 12.8V14a2 2 0 0 0 2 2h.6"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>"
        Case "volume"
            p = p & "<path d=""M4 9.5h3.2L12 5.4v13.2l-4.8-4.1H4z"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linejoin=""round""/>" & _
                    "<path d=""M15.4 9.2a4 4 0 0 1 0 5.6M18.2 6.6a7.6 7.6 0 0 1 0 10.8"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round""/>"
        Case "volume-mute"
            p = p & "<path d=""M4 9.5h3.2L12 5.4v13.2l-4.8-4.1H4z"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linejoin=""round""/>" & _
                    "<path d=""m15.6 9.6 4.6 4.8M20.2 9.6l-4.6 4.8"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round""/>"
        Case "queue"
            p = p & "<path d=""M4 7h11M4 12h11M4 17h7"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>" & _
                    "<circle cx=""18.5"" cy=""16"" r=""2.6"" stroke=""currentColor"" stroke-width=""1.7""/>" & _
                    "<path d=""M21.1 16V9.4l-4.4 1.1"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "music"
            p = p & "<path d=""M9 18.2V6.4l10-2v11.4"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<circle cx=""6.6"" cy=""18.2"" r=""2.6"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<circle cx=""16.6"" cy=""15.8"" r=""2.6"" stroke=""currentColor"" stroke-width=""1.8""/>"
        Case "album"
            p = p & "<circle cx=""12"" cy=""12"" r=""8.6"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<circle cx=""12"" cy=""12"" r=""2.4"" stroke=""currentColor"" stroke-width=""1.8""/>"
        Case "artist"
            p = p & "<circle cx=""12"" cy=""8.4"" r=""3.8"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""M4.8 20.2a7.4 7.4 0 0 1 14.4 0"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>"
        Case "playlist"
            p = p & "<path d=""M4 6h9M4 11h9M4 16h6"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>" & _
                    "<path d=""M16.5 18.4V8.6l4.5-1v8.6"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<circle cx=""14.6"" cy=""18.4"" r=""2"" stroke=""currentColor"" stroke-width=""1.7""/>"
        Case "user"
            p = p & "<circle cx=""12"" cy=""8.5"" r=""3.6"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""M5.2 20a7 7 0 0 1 13.6 0"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>"
        Case "users"
            p = p & "<circle cx=""9.4"" cy=""8.6"" r=""3.4"" stroke=""currentColor"" stroke-width=""1.7""/>" & _
                    "<path d=""M3 19.4a6.6 6.6 0 0 1 12.8 0"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round""/>" & _
                    "<path d=""M16 5.6a3.4 3.4 0 0 1 0 6.6M18 19.4a6.4 6.4 0 0 0-1.6-4.3"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round""/>"
        Case "logout"
            p = p & "<path d=""M14.5 8V6.2a2 2 0 0 0-2-2H6.2a2 2 0 0 0-2 2v11.6a2 2 0 0 0 2 2h6.3a2 2 0 0 0 2-2V16"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<path d=""M10.4 12h9.4m0 0-2.8-2.8M19.8 12l-2.8 2.8"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "settings"
            p = p & "<circle cx=""12"" cy=""12"" r=""2.9"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""M12 3.4v2.2M12 18.4v2.2M4.6 12H2.4M21.6 12h-2.2M6.8 6.8 5.2 5.2M18.8 18.8l-1.6-1.6M17.2 6.8l1.6-1.6M5.2 18.8l1.6-1.6"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>"
        Case "edit"
            p = p & "<path d=""M4.5 19.5h3.2l9.6-9.6-3.2-3.2-9.6 9.6z"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linejoin=""round""/>" & _
                    "<path d=""m14.6 5.3 3.2 3.2 1.5-1.5a1.4 1.4 0 0 0 0-2l-1.2-1.2a1.4 1.4 0 0 0-2 0z"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linejoin=""round""/>"
        Case "trash"
            p = p & "<path d=""M4.8 7h14.4M9.4 7V5.4a1 1 0 0 1 1-1h3.2a1 1 0 0 1 1 1V7"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<path d=""M6.6 7l.9 12a1.4 1.4 0 0 0 1.4 1.3h6.2a1.4 1.4 0 0 0 1.4-1.3l.9-12"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<path d=""M10.4 11v5.6M13.6 11v5.6"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>"
        Case "chevron-left"
            p = p & "<path d=""m14.5 5.5-7 6.5 7 6.5"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "chevron-right"
            p = p & "<path d=""m9.5 5.5 7 6.5-7 6.5"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "chevron-down"
            p = p & "<path d=""m5.5 9.5 6.5 7 6.5-7"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "more"
            p = p & "<circle cx=""5.5"" cy=""12"" r=""1.8"" fill=""currentColor""/><circle cx=""12"" cy=""12"" r=""1.8"" fill=""currentColor""/><circle cx=""18.5"" cy=""12"" r=""1.8"" fill=""currentColor""/>"
        Case "check"
            p = p & "<path d=""m5 12.8 4.6 4.4L19 6.8"" stroke=""currentColor"" stroke-width=""2.1"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "close"
            p = p & "<path d=""M6 6l12 12M18 6 6 18"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round""/>"
        Case "upload"
            p = p & "<path d=""M12 16.5V4.8m0 0L7.6 9.2M12 4.8l4.4 4.4"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<path d=""M4.6 15.4v2.8a1.6 1.6 0 0 0 1.6 1.6h11.6a1.6 1.6 0 0 0 1.6-1.6v-2.8"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round""/>"
        Case "eye"
            p = p & "<path d=""M2.6 12S6.2 6.4 12 6.4 21.4 12 21.4 12 17.8 17.6 12 17.6 2.6 12 2.6 12Z"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linejoin=""round""/>" & _
                    "<circle cx=""12"" cy=""12"" r=""2.9"" stroke=""currentColor"" stroke-width=""1.8""/>"
        Case "lock"
            p = p & "<rect x=""5"" y=""10.4"" width=""14"" height=""9.6"" rx=""1.8"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""M8.4 10.4V8a3.6 3.6 0 0 1 7.2 0v2.4"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>"
        Case "globe"
            p = p & "<circle cx=""12"" cy=""12"" r=""8.6"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""M3.4 12h17.2M12 3.4c2.4 2.4 2.4 14.8 0 17.2-2.4-2.4-2.4-14.8 0-17.2Z"" stroke=""currentColor"" stroke-width=""1.8""/>"
        Case "dashboard"
            p = p & "<rect x=""3.6"" y=""3.6"" width=""7"" height=""7"" rx=""1.4"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<rect x=""13.4"" y=""3.6"" width=""7"" height=""7"" rx=""1.4"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<rect x=""3.6"" y=""13.4"" width=""7"" height=""7"" rx=""1.4"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<rect x=""13.4"" y=""13.4"" width=""7"" height=""7"" rx=""1.4"" stroke=""currentColor"" stroke-width=""1.8""/>"
        Case "chart"
            p = p & "<path d=""M4 20V10.5M10 20V4.5M16 20v-7M22 20H2"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round""/>"
        Case "clock"
            p = p & "<circle cx=""12"" cy=""12"" r=""8.6"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""M12 7.4V12l3.4 2"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "calendar"
            p = p & "<rect x=""3.8"" y=""5.4"" width=""16.4"" height=""14.2"" rx=""2"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""M3.8 10h16.4M8.4 3.6v3.6M15.6 3.6v3.6"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round""/>"
        Case "mail"
            p = p & "<rect x=""3.4"" y=""5.6"" width=""17.2"" height=""12.8"" rx=""2"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""m4.6 7.4 7.4 5.4 7.4-5.4"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "shield"
            p = p & "<path d=""M12 3.4 19 6v5.6c0 4.2-2.9 7.5-7 9-4.1-1.5-7-4.8-7-9V6z"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linejoin=""round""/>" & _
                    "<path d=""m9 12 2.2 2.2L15.4 10"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "info"
            p = p & "<circle cx=""12"" cy=""12"" r=""8.6"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<path d=""M12 11v5.4M12 7.6h.01"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round""/>"
        Case "warning"
            p = p & "<path d=""M12 4.2 21 19.4H3z"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linejoin=""round""/>" & _
                    "<path d=""M12 9.6v4.2M12 16.6h.01"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round""/>"
        Case "trend"
            p = p & "<path d=""M4 16.4 9.4 11l3.4 3.4L20 7.4"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<path d=""M15.4 7.4H20v4.6"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round"" stroke-linejoin=""round""/>"
        Case "speaker"
            p = p & "<rect x=""6"" y=""3.6"" width=""12"" height=""16.8"" rx=""2"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<circle cx=""12"" cy=""15.4"" r=""2.8"" stroke=""currentColor"" stroke-width=""1.8""/>" & _
                    "<circle cx=""12"" cy=""7.8"" r=""1.2"" fill=""currentColor""/>"
        Case Else
            p = p & "<path d=""M9 18.2V6.4l10-2v11.4"" stroke=""currentColor"" stroke-width=""1.8"" stroke-linecap=""round"" stroke-linejoin=""round""/>" & _
                    "<circle cx=""6.6"" cy=""18.2"" r=""2.6"" stroke=""currentColor"" stroke-width=""1.8""/>"
    End Select

    Ico = p & "</svg>"
End Function
%>
