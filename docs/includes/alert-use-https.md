> [!NOTE]
>
> - Office Add-ins should use HTTPS, not HTTP, even while you're developing. If you're prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides. You may also have to run your command prompt or terminal as an administrator for the changes to be made.
>
> - If this is your first time developing an Office Add-in on your machine, you may be prompted in the command line to grant Microsoft Edge WebView a loopback exemption. When prompted, enter `Y` to allow the exemption. Note that you'll need administrator privileges to allow the exemption. Once allowed, you shouldn't be prompted for an exemption when you sideload Office Add-ins in the future (unless you remove the exemption from your machine). To learn more, see ["We can't open this add-in from localhost" when loading an Office Add-in or using Fiddler](/office/troubleshoot/office-suite-issues/cannot-open-add-in-from-localhost).
>
>    :::image type="content" source="../images/office-loopback-exemption.png" alt-text="The prompt in the command line to allow Microsoft Edge WebView a loopback exemption.":::
