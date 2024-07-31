> [!NOTE]
> Outlook on Windows: If you're running your add-in from localhost and see the error "We're sorry, we couldn't access *{your-add-in-name-here}*. Make sure you have a network connection. If the problem continues, please try again later.", you may need to enable a loopback exemption.
>
> 1. Close Outlook.
> 1. Open the **Task Manager** and ensure that the **msoadfsb.exe** process isn't running.
> 1. Set the [loopback exemption](/previous-versions/windows/apps/hh780593(v=win.10)?redirectedfrom=MSDN) in an elevated prompt.
>     - If you're using `https://localhost` and port 3000 (the default configuration), run the following command.
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_https___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>
>     - If you're using `http://localhost` and port 3000, run the following command.
>
>        ```command&nbsp;line
>        call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
>        ```
>
>      **Note**: If you're not using the default port 3000, replace it in the command with your actual port number.
> 1. Restart Outlook.
