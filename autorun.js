Office.onReady(() => {
});

function inserisciFirma(event) {
    const emailUtente = Office.context.mailbox.userProfile.emailAddress.toLowerCase();
    
    let firmaHTML = "";

    if (emailUtente === "manola.rossi@apes.it") {
        firmaHTML = `
<br><br>
<p style="font-family: Book Antiqua, serif; font-style: italic; color: #1f497d; font-size: 13px;">Cordiali saluti | Best Regards | Cordiales Saludos | Melhores Cumprimentos</p> <br>
<a href="https://www.apes.it" target="_blank" style="text-decoration: none;">
<img src="https://github.com/Helios-bar/email-signatures/blob/main/logo_manola.gif?raw=true" width="450" height="100" alt="Manola Rossi">
</a>
        `;
    } else if (emailUtente === "commerciale@apes.it") {
        firmaHTML = `
<br><br>
<p style="font-family: Book Antiqua, serif; font-style: italic; color: #1f497d; font-size: 13px;">Cordiali saluti | Best Regards | Cordiales Saludos | Melhores Cumprimentos</p> <br>
<a href="https://www.apes.it" target="_blank" style="text-decoration: none;">
<img src="https://github.com/Helios-bar/email-signatures/blob/main/logo_ingrid.gif?raw=true" width="450" height="100" alt="Ingrid Salvato">
</a>
        `;
    
    } else if (emailUtente === "mirco.rossi@apes.it") {
        firmaHTML = `
<br><br>
<p style="font-family: Book Antiqua, serif; font-style: italic; color: #1f497d; font-size: 13px;">Cordiali saluti | Best Regards | Cordiales Saludos | Melhores Cumprimentos</p> <br>
<a href="https://www.apes.it" target="_blank" style="text-decoration: none;">
<img src="https://github.com/Helios-bar/email-signatures/blob/main/logo_mirco.gif?raw=true" width="450" height="100" alt="Mirco Rossi">
</a>
        `;
    
    } else if (emailUtente === "matteo.bortoletto@apes.it") {
        firmaHTML = `
<br><br>
<p style="font-family: Book Antiqua, serif; font-style: italic; color: #1f497d; font-size: 13px;">Cordiali saluti | Best Regards | Cordiales Saludos | Melhores Cumprimentos</p> <br>
<a href="https://www.apes.it" target="_blank" style="text-decoration: none;">
<img src="https://github.com/Helios-bar/email-signatures/blob/main/logo_matteo.gif?raw=true" width="450" height="100" alt="Matteo Bortoletto">
</a>
        `;
    
    } else if (emailUtente === "elia.bottazzo@apes.it") {
        firmaHTML = `
<br><br>
<p style="font-family: Book Antiqua, serif; font-style: italic; color: #1f497d; font-size: 13px;">Cordiali saluti | Best Regards | Cordiales Saludos | Melhores Cumprimentos</p> <br>
<a href="https://www.apes.it" target="_blank" style="text-decoration: none;">
<img src="https://github.com/Helios-bar/email-signatures/blob/main/logo_elia.gif?raw=true" width="450" height="100" alt="Elia Bottazzo">
</a>
        `;

    } else if (emailUtente === "simone.rossi@apes.it") {
        firmaHTML = `
<br><br>
<p style="font-family: Book Antiqua, serif; font-style: italic; color: #1f497d; font-size: 13px;">Cordiali saluti | Best Regards | Cordiales Saludos | Melhores Cumprimentos</p> <br>
<a href="https://www.apes.it" target="_blank" style="text-decoration: none;">
<img src="https://github.com/Helios-bar/email-signatures/blob/main/logo-simone.gif?raw=true" width="450" height="100" alt="Simone Rossi">
</a>
        `;
    }

    // Incolla la firma se ne ha trovata una
    if (firmaHTML !== "") {
        Office.context.mailbox.item.body.setSignatureAsync(
            firmaHTML,
            { coercionType: Office.CoercionType.Html },
            function (asyncResult) {
                event.completed(); // Chiude il processo
            }
        );
    } else {
        event.completed();
    }
}

// Associa la funzione all'evento di Outlook
Office.actions.associate("inserisciFirma", inserisciFirma);