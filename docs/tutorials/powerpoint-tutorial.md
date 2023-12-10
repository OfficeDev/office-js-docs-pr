---
title: PowerPoint add-in tutorial
description: In this tutorial, you will build a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides.
ms.date: 12/11/2023
ms.service: powerpoint
#Customer intent: As a developer, I want to build a PowerPoint add-in that can interact with content in a PowerPoint document.
ms.localizationpriority: high
---

# Tutorial: Create a PowerPoint task pane add-in

In this tutorial, you'll use Visual Studio Code (VS Code), Visual Studio, or your preferred code editor to create a PowerPoint task pane add-in that:

> [!div class="checklist"]
>
> - Adds an image to a slide
> - Adds text to a slide
> - Gets slide metadata
> - Adds new slides
> - Navigates between slides

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# [Yeoman generator](#tab/yeomangenerator)

## Create the add-in

> [!TIP]
> If you've already completed the [Build your first PowerPoint task pane add-in](../quickstarts/powerpoint-quickstart.md?tabs=yeomangenerator) quick start using the Yeoman generator, and want to use that project as a starting point for this tutorial, go directly to the [Insert an image](#insert-an-image) section to start this tutorial.
>
> If you want a completed version of this tutorial, head over to the [Office Add-ins samples repo on GitHub](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/powerpoint-tutorial-yo).

### Prerequisites

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### Create the add-in project

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Choose a project type:** `Office Add-in Task Pane project`
- **Choose a script type:** `JavaScript`
- **What do you want to name your add-in?** `My Office Add-in`
- **Which Office client application would you like to support?** `PowerPoint`

![Prompts and answers for the Yeoman generator in a command line interface.](../images/yo-office-powerpoint.png)

After you complete the wizard, the generator creates the project and installs supporting Node components.

[!include[Node.js version 20 warning](../includes/node-20-warning-note.md)]

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### Complete setup

1. Navigate to the root directory of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Open your project in VS Code or your preferred code editor.

    [!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

## Insert an image

Complete the following steps to add code that inserts an image into a slide.

1. Open the project in your code editor.

1. In the root of the project, create a new file named **base64Image.js**.

1. Open the file **base64Image.js** and add the following code to specify the Base64-encoded string that represents an image.

    ```js
    export const base64Image =
        "iVBORw0KGgoAAAANSUhEUgAAAZAAAAEFCAIAAABCdiZrAAAACXBIWXMAAAsSAAALEgHS3X78AAAgAElEQVR42u2dzW9bV3rGn0w5wLBTRpSACAUDmDRowGoj1DdAtBA6suksZmtmV3Qj+i8w3XUB00X3pv8CX68Gswq96aKLhI5bCKiM+gpVphIa1qQBcQbyQB/hTJlpOHUXlyEvD885vLxfvCSfH7KIJVuUrnif+z7nPOd933v37h0IIWQe+BEvASGEgkUIIRQsQggFixBCKFiEEELBIoRQsAghhIJFCCEULEIIBYsQQihYhBBCwSKEULAIIYSCRQghFCxCCAWLEEIoWIQQQsEihCwQCV4CEgDdJvYM9C77f9x8gkyJV4UEznvs6U780rvAfgGdg5EPbr9CyuC1IbSEJGa8KopqBWC/gI7Fa0MoWCROHJZw/lxWdl3isITeBa8QoWCRyOk2JR9sVdF+qvwnnQPsF+SaRSEjFCwSCr0LNCo4rYkfb5s4vj/h33YOcFSWy59VlIsgIRQs4pHTGvYMdJvIjupOx5Ir0Tjtp5K/mTKwXsSLq2hUWG0R93CXkKg9oL0+ldnFpil+yhlicIM06NA2cXgXySyuV7Fe5CUnFCziyQO2qmg8BIDUDWzVkUiPfHY8xOCGT77EWkH84FEZbx4DwOotbJpI5nj5CQWLTOMBj8votuRqBWDP8KJWABIr2KpLwlmHpeHKff4BsmXxFQmhYBGlBxzoy7YlljxOcfFAMottS6JH+4Xh69IhEgoWcesBNdVQozLyd7whrdrGbSYdIqFgkQkecMD4epO9QB4I46v4tmbtGeK3QYdIKFhE7gEHjO/odSzsfRzkS1+5h42q+MGOhf2CuPlIh0goWPSAogcccP2RJHI1riP+kQYdVK9Fh0goWPSAk82a5xCDG4zPJaWTxnvSIVKwKFj0gEq1go8QgxtUQQeNZtEhUrB4FZbaA9pIN+98hhhcatbNpqRoGgRKpdAhUrDIMnpAjVrpJSNApK/uRi7pEClYZIk84KDGGQ+IBhhicMP6HRg1ycedgVI6RELBWl4POFCr8VWkszpe3o76G1aFs9ws+dMhUrDIInvAAeMB0ZBCDG6QBh2kgVI6RAoWWRYPqBEI9+oQEtKgg3sNpUOkYJGF8oADxgOioUauXKIKOkxV99EhUrDIgnhAG+mCUQQhBpeaNb4JgOn3AegQKVhkvj2gjXRLLrIQgxtUQYdpNYsOkYJF5tUDarQg4hCDS1u3VZd83IOw0iFSsMiceUCNWp3WYH0Wx59R6ls9W1c6RAoWmQ8PaCNdz55hiMEN4zsDNhMDpXSIFCwylx5Qo1a9C3yVi69a2ajCWZ43NOkQKVgkph5wwHi+KQ4hBs9SC9+RMTpEChaJlwfUFylWEafP5uMKqIIOPv0sHSIFi8TFAzpLiXxF/KCbdetEGutFUSa6TXQsdKypv42UgZQhfrWOhbO6q8nPqqCD/zU4OkQKFpm9B7SRbrTpQwzJHNaL/VHyiRVF0dfC2xpOzMnKlUgjW0amhGRW/ZM+w5sqzuqTNWtb9nKBZDLoEClYZGYe0EYaENWHGDaquHJv5CPnz/H9BToWkjmsFkTdOX0GS22p1ovYNEdUr9vCeR3dJlIG1gojn2o8RKPiRX+D0iw6RAoWmYEH1HioiQZqq47VW32dalUlfi1fQf7ByEdUQpMpYfOJ46UPcFweKaMSaWyaWL8z/Mibxzgqe3G4CC6pT4dIwSLReUCNWrkJMdjh8sMSuk1d3bReRGb3hy97iS/SEl+5bQ0LqM4B9gvytaptC6kbwz++vD3ZG0r3EBDoWUg6RAoWCd0D9isXReTKTYghZbhdUB/UYlKV2TSHitZtYc9QrqynDGy/GnGg+4XJr779ShJ0gNdAKR3i/PAjXoIZe8BGBS+uhqtWAF4VXUWu3G//ORVqdVRiEumhWgFoVHT7gB1LnFAvVaJxYZJ+qx/XRuo1X0+RFqzPsF/QFZuEgrVcHnDPCGbFylnajN/wAZZvqgpR8IzO275tTvjnwl/4sORC6C9xWJLoYCKNrbpuR3Jazp/jxdUJmksoWIvvAfcLsD4LuLfn5hOJhWlVQ+lyNZDFcUl636GY5/Wpyzo3FRZ+WBeT1JhpGDVlIMMbjYfYM3Ba4zuXgkUPGBD5B5Kl6LaJ4/uh/CCDTvDjW4ROxZm4gj7+dwZLY24067AkF9OtesCaRYdIwaIHDIzMrmSzv2NNTgl4fLlSXw6kjs8pWN+FfHu3n8p/xpSBjWrwL0eHSMGiB/TL+h1JnNJ+xTA6MawXh1ogTWA5S5tvLS8vMVUM6s1j+TKZEASjQ6RgkVl6wH4pcUM+zs8qBq9WyRyMGozP+5J0/nzygrrLSkS4ONPmNg/vyr1npiQG9+kQKVhkBh5woFbSI8EuQwxTkS1j2xoG0zsHeBVcRsl/RNMqyoMOG9WRjAUd4pzD4GhoHjDsMIEqchX48JuUgU1zJN+kSa4D+LnjHfXiqqsa5Oejb8J/fs9TAZjFtiXXvgADpaqXZsqUFRY94NRq1agErFbrRWzVR9Tq9JlOrWy75NncCf982n+o+sYCDJTSIVKw6AGnRhoQbZsBv3S+MlyxAtC7xPF9WMUJDsi5M+gmVCWImpvolorOgXzTMPBAKR0iBWvuPWB4+4CiWj2Rz3MPcFSXHb90NmawbWDLRVZAc2pHZTkF2fWDKugQRqBUCvcQKVj0gI6qRxYQtfvGBIUdvHQ2fmk/VR7fk5Q5jr+2fmfygrpTfM+fu8qa6lEFHcIIlGocolWkQwwcLrr79oBB9YRxg7SDXbDjJISue71LHJWnrno+vRh+BX2Xq2QOO6+Hf3TTXsYl43M3BhVcZFNjEyvIluUNvAgrrIX1gINqRdpvM0C1EhatbBvowaM5neOVe/L2VX176/jip88CUysAhyV5SRheoFRSfV+i8RAvckH+XKyweBW8qNWeEelEP1XkKqgQw3j/T3sxyNv6cSKNm02xA3KrOvLV1gq4Xh1u3vUusWcE7KESK7jZlHvSoDqU+q/4CAUrItomWtUoRvup1KpRCWxb0KiNqFXvcoreWCem/ETh+ILRYJnvJzlxz+7wrt/l9qkuHUIIrMk9bxaZEjIltl2mYMWDjoVWFae1sAouVeQq2LUYZwfRaVG1dR9PnKp802EpxG016TCOgZsOb6tk9RayZVZVFKwZ8cff4b/+Htcq8sd17wInJt5UA17SUqnVWR0vbwf5Qn5KgPO6bo0mU0K2LJetbgtvqjgxQw8uqcbthDH+OrHS/5FV19MuJDXreoSCFQC9C3yxisQK8hVk1dteZ3W8qQY2VFm68OF/emj0JNJ430DKQCKN3gU6FrrNSHf9VaMrfI68F+ynXVKpkhxndRyX0TlQzv4hFKyABWuwMPGROWxiJ6kdmmibaJu+7gTpPRbgDbZsqJa9/T8AMrvIlnWx/m4Tx+XhY4yC5RXGGjzRbeHlbd3ZsWQO+Qp2mth84nFtSBoQtS0M1cobqqCD50BpMovrj/Dpufyk1OBXZueKgyq6KVjEI/bZMf3ef6aErTp2XiOzO8UtIe0gCuCoHMWm5MLWyJfK09HTdihdvwPjc+w0J4wvbJv4KhfF2VIKFnHLm8f4KjfhkF0yh00TN5vYfDJ510wVED0qR7ENv7Sa5SZQmlhB/gF2XsOoTdj+O6tjz8Dh3Tlbaow9XMNy/153rGGpDIJ+Ycv5bm6bcvVR5YaiPFCy8Kze6s+4lj4VpIHS1Vv4sORqa09YrlL5fa5hUbBmLFiDd/am6Soi0LtAqzqyMK9Sq8BDDEQVdMBooDSxgvXihAV14RfqxgBSsChYcREsmyv3lImtcU5raJs4q8sjV/MYYpgLrj9SxlP2C/iuiXxFl1EYL4GPym5/TRQsCla8BKu/3qFNbLl80a9yVKuwUIWzpmKQrnIPBcsrXHQPT+AucXzf70l91lahclT2FV7tNmEV8fI2t24jI8FLEC52Ysv9wpbAtsVLGNNy2+VyFWGFNX+4SWyReYHpKgrWUuAmsUXiDNNVFKwlsxJBLGyRGVh7LlfFAq5hzeTd38LL27oo0ABpnykSIG766pzWYH3GS0XBWvJr7yLg8/1F1J18l4pk1lXuhM1CaQkJPixN/jvXKlGMpVpa8u7CvSkj9CGshIIV92e7tOvxeBXGhGFIrN6Sp0ZPa5Jw1gfsdEzBWmbGb4BuE4d3JbdKtszHe1jllZTjsqTBvJtymFCwFpbxpRM77nAouzE+MnnBAiazK++rYZ9Flw4B4mODgrWkpG5I1nHf1gDFrPa1gveRNmQc+5jnOL2L/pDqzoGkN2mArpChFgrWXD3eS5J38KDJjDTKsMG4aaDlrXTjr1UdJkJPTLpCChYBAEmzSqcHOX8utySZXV65AFBFGezjgULBS1dIwaIflDzehVVeVZHFiIN/VFEGoZtVtyUxbtwrpGDNDb3fheUH26Z4Nq3bkhw5TKT9dtciqihDtynpWN2mK6RgzS/vemH5QemU9kZF0tohX6Er8VteSTmWPQlOZa5w4gwRQsFaZD/Yu5APLOhdyvs6XOfqu+faVhFlOKsrfwXjRRZHzFOwlumeKbkqr2xaVUmOdL3IiEPA5ZXmhPn4b2edy1gUrOVh/O2uaY/Vu2TEITi1eiCPMrRNnD9XC9Yz0Zgnc3SFFKxl9YPd5oT+Su2nkgQjIw7TklhR7ldMbOBzQldIwVpOxu+Z8SWScY7K8iKLEQf3bFTlUYZWdZjXVT4zTLrCGD16eAlm6QfdCJZ9WEdYLbYjDmG3FU/mRqoJD90EV3+Ga//o5aUPS77m2QiFrbQm6l24+ok6B+g2R0pj2xWy9SgFa6HV6o74kO9Ykx/vNsdlyficfGVkanRIgpV/4Euw3v/E4xZBMheYYKn2VZ0HcfS0quK6YaaE4/t8U9MSLlN55X4aRedAXouxVZab54Q0ytBtTnH933KvkIJFwdIEGsaRVjeZEiMOHsurRmWKyTfdlrj1wb1CCtZy+cHT2nSjorotuWbFvMj6w6/xhxN81xL/G/zsvY7ks384wfdBDHBURRmkB3EmukIBHpOaBVzDmlF55Wa5ffyeyZZF4VsrILM79e0XGb/5JX7zS8nHt+r92rDz79gvhPPWVkcZpF0S9cgTpHf51maFtQSCpTqOo0d1WCfPQRUyVFGGs7ouKaq5+IJmJdJYv8PLTMFaDj/ojcZDyd5ZMkd7IqKKMsDHqEcGsihYS+oHT0zvX016v3FQhYBqrV1/EGeCKxw7pkPBomAtGokV8W3dbXq/Z6A4rMNpYE5Wb8mjDPA9SZuucOb3Ey9B6OVVUH5wwFEZW3Xxg5kSTkxfUmjj/MrCdz7+ovpvclxYo2HTVKqVz5xtqyo6zfWil+VIQsGaGz/4xnevBelhHQD5Cl7eDqA88fCpcX6cns0Fv3JPHmUQWrZ7Y/yYDvcKaQkX2Q+6P46j5+uS5IN2xCEO9C7xrTWbC36toiyOpgq+KS25SVfICmtpyqsTM5ivbA/7HN8Iy1emjqQKOGu0lIHrj+SfEhD+5mFJ0t85AlQDJrrNwA6Kt01xuZCukIK1sILlIS+qolGRLJDZEQc/N6dmxqfmU85dufbTANbpPKCa3wXfa+3Co6JjIWX4coWzWt2jJSRT+EGftc/4nSNdlMmWo86R5ivDg3XdlryBVwR8ZCrVIdiTACdjrnBaJx7g24CCRcIqrwKvO1pVifNKpCPtoZwyRlrQfD0jM6iJMgQuoEyQUrAWX7B6F8ELVu8S38jMTqYUXS8BZ4ag8VBnGyP7NgQb6z/qMX7ZhV/lepGnoyhYMeP/vouRHxzw5rG80V0008CcZrBzEORS0VSoogxQDBz0D6fpULAWSrAi8IPDukYmE2uF0LfbBTPooQVCIGiiDG0zrEbG7ac8pkPBWiCEwEG3GeLOd/up3IiFXWQ5Xdjx/ZntfKmiDEC4FR9dIQVrQUhmxQXgsLf5pXem0JE9PDN4/jyAELnnS62JMoTa8P7EpCukYC0EH4QZv5JiH9YZJ6SIg9MM9i5nZgY1VWQgB3EmXnNh9ZCCRcGaSz4cvYE7VhQjoaSHdUKKODjNYIDzuKZl9ZZSI76pRJF1oiukYC2CH3TGoBHccRw99mGdcQKPODjN4Omz2YTabVRa3G3izeMovoHxc+wssihYc+8H30Z1Szcq8tBmgKvv8TGDmV3xweC8DtEwPk2HgkXBmm8/eFoLd+lXuH+kCzcBRhycZtAqzibUDiCxoiyvzuqRjuQQyuf1Ilu/UrDm2Q9G7Jikh3WCKrKcZvDN41BC7X/+NzBq+Nk3yurJZnx6UPTllap8/oBFFgVrfv1gxILVu5QfnUvmcOWe3y8+CBB0DuRHgvyI1F//Cp9+i7/6Bdbv4E/zuv5/yayyH3QYB3EmVrXCr/jDEu8DCtZ8+sG2OYNz+e2n8m27a76ngQ3+eYDtrlZv9UXqp3+BRMrVP9FUi1/PQiwEwUoZdIUULPrBaZAeoAtqUEXj4SzbOWmiDG0zuuVC4bcsyDddIQVrDhCO43iblhrMLfRMmSP1+fCP4ITz//4WHUuZ7dpQJ0VndfR6vHkDXSEFa/4E68Sc5Tejuns/Mn3dmVY4tUOvg9//J379C/zbTdQ/wN7HcsHSRBla1dmUV3SFFKy5JHVD7HAS9nEcPefP5YZ0rTDd8BtBBIMKtf/oJwDwP/+N869w/Hf44n3861/iP/4WFy+U/0QTZfB/EGe9qOyo5bKkFa4MXWE4sKd7OOVVtxnFcRw9x2X5cs+miRdXXX2Fb62RwRMB5hga/4Df/2o6+dNEGfwfxLle7ddEnqOwp7WRY9gfliJK27PCIh4f0YJDmTmqwzruIw69C5zVh/8FyG//aTq10nRl8H8QJ1/pq1VmVzKIyCXCpaYrpGDNkx98W4vFN3ZUlucPrlXm7JhueE2vEukRKfS8kdo5EDdPPWsfoWBF6gfP6gEvAKcM5Cv9/zIl5a0rKZEu5bVeUBGHaFi9pbz5/R/E2aiOaHcy611oTkwKVti89+7dO14Fd49QC3sfyz+183qkwjosBXacba2AfEVcJrdlSHUKR9SmFdxsyjXuRW6WO2vu+eRL5USc/YKvaHvKwPYriZV+kfPy1ZJZ7Iz63D1DuZT5c953rLBi4gcDyYsmc9g08cmXkk29xAryD3CzqbyNBXVTzbnyE3GIrnrdVf6YpzW/B3Gc247dVl++PRdZ3Za40qf5OrM6N07Boh8U7yKfO1a2VO28njCeM7GCT750dWupDuv4iThEQ2JFZ119TsRZL478+F+Xhsthnv2ysPSu6TbzLYc/U7BmgvCm9Bm/ShnYtiRS1TlA4yEaD3H+fEQQN5+46imq2q3fqMb62mbLyvld/g/iOM8k2mcDBl/Tc5ElFNfJXHQDIilYxIVa3Rm5o3wex0kZ2KqL+3ftp3hxFXsGGhU0Ktgv4Is0Xt4eytaVe5MrAlXT95Qx9Zj1yNBEGXoXk+c5pwydZR5EGWzXPCjWfBZZvUvxicWldwrWbHjXm1xe+Vy92jRH1KpzgL2P5U3Tz+ojp2TyD5SVyADV9r+wTRYfNFGGVnWC706kYdTwyZfYqktkS4gytKrDKzxw9EEVWexBSsGaDb3fTRYsP3lRofl65wD7BV1fBGFH302RJbWrwt0bEzRRBjcHca79UECt3pLIllOju60RKXd+cW9F1umzkQV1ukIKVoz8oLME8Hkcx6l9vUvsFyZvJDnv29XC5JdQFVlOfxSf8krFUXlCeZXMiWLnlC3BBY+30BqUb56LrBO6QgpWHAUr0OV2Z49NVUJdoGMNb103iqNq+o7wx0RPV2yqowzd5uSMW7eJPUOymDiQLWc1NL6057/Icr9XSChY8ypYmnUQvWYNcBPLUk3WEfb4Z0ggUYZuE1YR1meSWmxgBp1r7SrF8VZkdQ5Glh2TubjHRyhYS+cHO5bfXXan9LhPFTrvBDfHiVWHdRCbiIMmynBWn24T9rSGr3LKo9HfXygX9Z11nLciS7jIbOlHwYpXeeW/PcP3DpHSz4xRlVQu+x84N8WcxCHikFjR7QB4OOdsByBe3pYsLyaz2H6FTVOuj4PX8lZkveVeIQUrzoI10cQl0hNaxDkrLDfbdon0yMKT+0Mqvcv4Rhw2qsqqx89BnLM69gx5CZzZxc5ryev6LLKEGauJdGCjISlYxK8fnHgcZ72Im01dh1+MtsfL7E7OVW1UR/bLT8wpvn/VYZ3ZRhxSN3S1jM+DOGuF4b6EcFoAwJV7uNkUk1+DqtlbkSUU3SyyKFhzU14Zn/crF826eO9iZP9r09S1kcmWR+zb6bOpl/xVh3VmGHHQ7FT6b9k+qJJ6l3hVxJ4h7jYOjpQPtKljDWs6D0UWE6QUrFiQWBl53gpCI7d7Pyyg6B/UDUer39Vb2KpLNCuRxkYV1x+NfHEPjX1Vh3Uwo4jD+h2lmvufiOM85m235ek2cVjCy9uizUysYPMJdn6QLT8rWcI0HbpCCtZ8lFdOd5C6oSuy7LvIaZGcD/y1AjIlbFsjDY57l97HmqpM1kwiDvryymcDDLuNcrclbpKe1bFfwOFd8esns9h80k9s+SmyGMgKGjbwc81ZvT+Rwfh85J3npodcIo2bzb4rPH+O/cIEQRQOFWqe4frjOxPZfCIvHAY/bDTkHyjlwE6BBjVAO5nTLd7lH8i+gdbQIx/endp6f3o+LJN7F/hitf//mq6EhBVWkH7QqVbdpqutK2d4WjO7eFCyfZVD4+GEgz7+1QrqoMBaIbqIw8QoQ1BqBXXyw3adL65KfpvOFT2fK1l0hRSsOfCD475m05zwdLXvnz0DL66i8VByx3YOsGcEMDJeOPo7UvVENahCE2VwcxAnQLpN7Bfw8rZygd/DShb3CilYMRKsN67Xp3sXw/Upu1mopn2KfXzXqGHnNfIPROGwTWVQM01VveGTuSgiDvoog+cpgT69/4scju8HU9kJx3TWi3M2ryhmcA1rmvexVcSnjntbM5ZCxaY5YrXsjaSOhY6FRBopA8kcUoauIUnjod8tM0kxpVhC6l0o85ZBoVnKiXgdTeJV09iojvy+vM2nEC6vPaOEa1gUrNAFq22OpNWPyl5GeAqa5Z7z52hUAh5oOkAY/DOgbeLwbmjl6h0Yak/tcyJOYDWggY1qf9vUw6I7xqbpnNZgfUbBoiWM3A96a89wWJrabpw+w8vb2C+EpVZQr75nSiFGHDRRhrYZC7Wy6+j9AqzPvKRzB3WZc7WRrpAVVhRc/AvSPxOfk37sxnoRawUkc0ikJR6w28J5HWd1nNYiGgm1/Up+cigka3blnq4/xLzMTPT2wx6WkCmxwqJghcnvj/DTDXElItgVk/cNAPjWms3QOjtbr6oKA/5h1eNdAbSqOL6/UG+exMrI6udpDYk0BYuCFSZ//B3+5M/6/9+7wFe5IPNBMUG1sBJsehPA9Ue6iTgLeW2FvHHHcttEiDjgGpZrBmqFIKalxhPVYZ1gIw6a+V0I4iBOPBEie1QrCtbM3nwLQ+dAua6cLQfWxeEjU/mpbhONh4t5bdtPOZ6egjULuk1f01JjjqrpeyLtfYC7k9VburWbwCNmfM5RsFheLbQcqyfrCJMTvaFpu9qxIj2IEz0nJu8eClb0tf2iv+1Uh3Xgu1XWlXu6TqpH5QW/sOfPAztQRcEiruhYvqalzgW9S3yjsGZrBe/9BhIruKZ2fGf1uCRFWZ5TsFjVzxlvHitrAc9FluawN3y3bGd5TsEiEt4uzRNStf6dzMkb3enRRxna5uLXrf0K/SCApkAULOK2nl+k8yITaoGnyqOL2fLUp+E+Mr2II4t0QsHyJVhLhUpH7L4r7pkYZViex8BSFekULApWpGgm60wVcdCom7N59JLQbXHp3TMJXgK3vOvBqKF3gY6FbhPdJr5rLn5p8HVppJeTk+tVV10c9ONjF/UgzshNtoKUgR+nkTKGbRqJJ3j42f8Ds4luEx2rr2XfX6BjLdRNqJqsA8AqTgj967sydJt4cXWh3gypG8M2DKsFAGzJQMGaE2wzdV7v/3/vYl43wpJZbFty0ZmoOJr5XQiha02U1+QnOSRz/ZbWdmsgTWiDULDmkt5Fv93VfPlKje40KsrjykJr4HFBn23Lds9ujoaOgkVfGWtfqXF2mvZVQgcogZi0bKebo2CRBfSVmo7G0gahmv6lsy2v6OYoWMuL7ewiftPPyleqJutA1oJd1SFe9fcXz83ZD5vvmlPPXiUUrBBpm8Pooz1gZmAr7LtlYXylZiqXUDFldnVtZAIfHTZbN6e67IkVZMvIllm+UbDiR6uKRkWuDs5HfTI39CPz6Cs10/QGa1L6KIOf4ayzdXNTFbaZXWxUKVUUrBhjh7bdJyHt289pW+LvKzUrU4OIgz7KoNlVjJub8ybxmV3kK9xJpGDNj2wdlX3Fi2LuKzV7f0dlvK3pogzjW4rxdHOef3H5CvcWKVhzSLeJ43KQrd/j4yuTOeUqsl21ae7YjoXT2tyUk1N51Y9MShUFa845q6NRCTdtNFtfGc9rjgiDIMks8hXuA1KwFojTGo7LUcfZZ+srI3Nz3/3g6aKP2nITkIK1yLRNHJVnHF6fua/06eZsVYrDYaYr93CtQqmiYC00024jRkZMfKUtSQM3B8RxLAU3ASlYSydb31Tw5vEcfKsh+cqZuznPV2OjyhHzFKylpNtEozKXzVXc+8p4ujkPpG7gepWbgBSspSeCbcRoGA+LzkX3GDdmmZuAsXpc8hLMkrUC1uo4q+Pr0nINYpiLQjJb1kX2ySzgEIp4yNZOE5tPkMzyYsSlYLzZpFpRsIiaTAnbFvIPph75R4L8Lexi5/WEIdWEgkUAIJFGvoKbTS+jlYlPVm9h5zU2TUYWKFhketnaeY3MLi9GRFL1yZfYqlOqKFjEK8kcNk1sv+qHoUgoFzmLzSfYqjOyQMEiQZAysFXHJ19OMWaZuCpjV3D9EXbYv5iCRQJnrYBti9uIgUmVvYzBIcUAAAIqSURBVAmYLfNiULBIaGRK2GlyG9HfNdzFtsVNQAoWiYrBNiJlayq4CUjBIjMyNWnkK9i2uI3oVqq4CUjBIjPG3kbcec1tRPUlysL4nJuAFCwSJ9mytxEpWyNF6Ao2n2CnqZyXQShYZGasFbBV5zZiX6rsTUDmFShYJNbY24jXHy3venxmt39omZuAFCwyH2TLy7iNuH6nvwlIqaJgkXmzRcu0jWhvAho1bgJSsMg8M9hGXL+zoD9gtp9X4CYgBYssjmwZtUXbRrQPLe80KVUULLKI2NuIxudzv41obwJuW9wEpGCRRWe92O/FPKfr8VfucROQgkWWjExp/rYR7c7FG1VKFQWLLB+DXszx30a0NwF5aJlQsChb/W3EeMpW6gY3AQkFi4xipx9itY1obwJuW5QqIj5keQkIEJuRrhxfSlhhkSlka4YjXTm+lFCwyNREP9KV40sJBYv4sGY/bCNeuRfuC63ewvYrbgISChYJQrY2qmFtIw46F6cMXmlCwSIBEfhIV44vJRQsEi6BjHTl+FJCwSLR4XmkK8eXEgoWmQ3TjnTl+FJCwSIzZjDSVQPHl5JAee/du3e8CsQX3Sa6Y730pB8khIJFCKElJIQQChYhhFCwCCEULEIIoWARQggFixBCwSKEEAoWIYRQsAghFCxCCKFgEUIIBYsQQsEihBAKFiGEULAIIRQsQgihYBFCCAWLEELBIoQQChYhhILFS0AIoWARQkjA/D87uqZQTj7xTgAAAABJRU5ErkJggg==";
    ```

1. Open the file **./src/taskpane/taskpane.html**. This file contains the HTML markup for the task pane.

1. Locate the `<body>` element. Replace it with the following markup, then save the file.

    ```html
    <body class="ms-font-m ms-welcome ms-Fabric">
        <!-- TODO2: Update the header node. -->
        <header class="ms-welcome__header ms-bgColor-neutralLighter">
            <img width="90" height="90" src="../../assets/logo-filled.png" alt="Contoso" title="Contoso" />
            <h1 class="ms-font-su">Welcome</h1>
        </header>
        <section id="sideload-msg" class="ms-welcome__main">
            <h2 class="ms-font-xl">Please <a target="_blank" href="https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing">sideload</a> your add-in to see app body.</h2>
        </section>
        <main id="app-body" class="ms-welcome__main" style="display: none;">
            <div class="padding">
                <!-- TODO1: Create the insert-image button. -->
                <!-- TODO3: Create the insert-text button. -->
                <!-- TODO4: Create the get-slide-metadata button. -->
                <!-- TODO5: Create the add-slides and go-to-slide buttons. -->
            </div>
        </main>
        <section id="display-msg" class="ms-welcome__main">
            <div class="padding">
                <h3>Message</h3>
                <div id="message"></div>
            </div>
        </section>
    </body>
    ```

1. In the **taskpane.html** file, replace `TODO1` with the following markup. This markup defines the **Insert Image** button that will appear within the add-in's task pane.

    ```html
    <button class="ms-Button" id="insert-image">Insert Image</button><br/><br/>
    ```

1. Open the file **./src/taskpane/taskpane.js**. This file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office client application. Replace the entire contents with the following code and save the file.

    ```js
    /*
     * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
     * See LICENSE in the project root for license information.
     */
    
    /* global document, Office */

    // TODO1: Import Base64-encoded string for image.
    Office.onReady((info) => {
      if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        // TODO2: Assign event handler for insert-image button.
        // TODO4: Assign event handler for insert-text button.
        // TODO6: Assign event handler for get-slide-metadata button.
        // TODO8: Assign event handlers for add-slides and the four navigation buttons.
      }
    });
    
    // TODO3: Define the insertImage function.

    // TODO5: Define the insertText function.

    // TODO7: Define the getSlideMetadata function.

    // TODO9: Define the addSlides and navigation functions.

    async function clearMessage(callback) {
      document.getElementById("message").innerText = "";
      await callback();
    }

    function setMessage(message) {
      document.getElementById("message").innerText = message;
    }

    // Default helper for invoking an action and handling errors.
    async function tryCatch(callback) {
      try {
        document.getElementById("message").innerText = "";
        await callback();
      } catch (error) {
        setMessage("Error: " + error.toString());
      }
    }
    ```

1. In the **taskpane.js** file above the `Office.onReady` function call near the top of the file, replace `TODO1` with the following code. This code imports the variable that you defined previously in the file **./base64Image.js**.

    ```js
    import { base64Image } from "../../base64Image";
    ```

1. In the **taskpane.js** file, replace `TODO2` with the following code to assign the event handler for the **Insert Image** button.

    ```js
    document.getElementById("insert-image").onclick = () => clearMessage(insertImage);
    ```

1. In the **taskpane.js** file, replace `TODO3` with the following code to define the `insertImage` function. This function uses the Office JavaScript API to insert the image into the document. Note:

    - The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsync` request indicates the type of data being inserted.

    - The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.

    ```js
    function insertImage() {
      // Call Office.js to insert the image into the document.
      Office.context.document.setSelectedDataAsync(
        base64Image,
        {
          coercionType: Office.CoercionType.Image
        },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            setMessage("Error: " + asyncResult.error.message);
          }
        }
      );
    }
    ```

1. Save all your changes to the project.

### Test the add-in

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. Complete the following steps to start the local web server and sideload your add-in.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.

        ```command&nbsp;line
        npm start
        ```

    - To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. If the add-in task pane isn't already open in PowerPoint, choose the **Home** tab, and then choose the **Show Taskpane** button on the ribbon to open the add-in task pane.

    ![The Show Taskpane button highlighted on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-yo-show-taskpane-button.png)

1. In the task pane, choose the **Insert Image** button to add the image to the current slide.

    ![The PowerPoint add-in with the Insert Image button highlighted.](../images/powerpoint-tutorial-yo-insert-image-button.png)

## Customize user interface (UI) elements

Complete the following steps to add markup that customizes the task pane UI.

1. In the **taskpane.html** file, replace `TODO2` and the current header section with the following markup to update the header section and title in the task pane. Note:

    - The styles that begin with `ms-` are defined by [Fabric Core in Office Add-ins](../design/fabric-core.md), a JavaScript front-end framework for building user experiences for Office. The **taskpane.html** file includes a reference to the Fabric Core stylesheet.

    ```html
    <header id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </header>
    ```

1. Save all your changes to the project.

### Test the add-in

1. If the local web server isn't already running, complete the following steps to start the local web server and sideload your add-in.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.

        ```command&nbsp;line
        npm start
        ```

    - To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button highlighted on the PowerPoint Home ribbon.](../images/powerpoint-tutorial-yo-show-taskpane-button.png)

1. Notice that the task pane now contains an updated header section and title.

    ![The PowerPoint add-in with Insert Image button.](../images/powerpoint-tutorial-yo-new-task-pane-ui.png)

## Insert text

Complete the following steps to add code that inserts text into the title slide which contains an image.

1. In the **taskpane.html** file, replace `TODO3` with the following markup. This markup defines the **Insert Text** button that will appear within the add-in's task pane.

    ```html
    <button class="ms-Button" id="insert-text">Insert Text</button><br/><br/>
    ```

1. In the **taskpane.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.

    ```js
    document.getElementById("insert-text").onclick = () => clearMessage(insertText);
    ```

1. In the **taskpane.js** file, replace `TODO5` with the following code to define the `insertText` function. This function inserts text into the current slide.

    ```js
    function insertText() {
      Office.context.document.setSelectedDataAsync("Hello World!", (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
        }
      });
    }
    ```

1. Save all your changes to the project.

### Test the add-in

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. If the local web server isn't already running, complete the following steps to start the local web server and sideload your add-in.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.

        ```command&nbsp;line
        npm start
        ```

    - To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-yo-show-taskpane-button.png)

1. In the task pane, choose the **Insert Image** button to add the image to the current slide, then choose a design for the slide that contains a text box for the title.

    ![The Insert Image button highlighted in the add-in.](../images/powerpoint-tutorial-yo-insert-image.png)

    ![The selected PowerPoint title slide highlighted in the add-in.](../images/powerpoint-tutorial-yo-insert-image-slide-design.png)

1. Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.

    ![The selected PowerPoint title slide with the Insert Text button highlighted in the add-in.](../images/powerpoint-tutorial-yo-insert-text.png)

## Get slide metadata

Complete the following steps to add code that retrieves metadata for the selected slide.

1. In the **taskpane.html** file, replace `TODO4` with the following markup. This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.

    ```html
    <button class="ms-Button" id="get-slide-metadata">Get Slide Metadata</button><br/><br/>
    ```

1. In the **taskpane.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.

    ```js
    document.getElementById("get-slide-metadata").onclick = () => clearMessage(getSlideMetadata);
    ```

1. In the **taskpane.js** file, replace `TODO7` with the following code to define the `getSlideMetadata` function. This function retrieves metadata for the selected slides and writes it to the Message section in the add-in task pane.

    ```js
    function getSlideMetadata() {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
        } else {
          setMessage("Metadata for selected slides: " + JSON.stringify(asyncResult.value));
        }
      });
    }
    ```

1. Save all your changes to the project.

### Test the add-in

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. If the local web server isn't already running, complete the following steps to start the local web server and sideload your add-in.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.

        ```command&nbsp;line
        npm start
        ```

    - To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button on the PowerPoint Home ribbon.](../images/powerpoint-tutorial-yo-show-taskpane-button.png)

1. In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written in the Message section below the buttons in the task pane. In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.

    ![The Get Slide Metadata button highlighted in the add-in.](../images/powerpoint-tutorial-yo-get-slide-metadata.png)

## Navigate between slides

Complete the following steps to add code that navigates between the slides of a document.

1. In the **taskpane.html** file, replace `TODO5` with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.

    ```html
    <button class="ms-Button" id="add-slides">Add Slides</button><br/><br/>
    <button class="ms-Button" id="go-to-first-slide">Go to First Slide</button><br/><br/>
    <button class="ms-Button" id="go-to-next-slide">Go to Next Slide</button><br/><br/>
    <button class="ms-Button" id="go-to-previous-slide">Go to Previous Slide</button><br/><br/>
    <button class="ms-Button" id="go-to-last-slide">Go to Last Slide</button><br/><br/>
    ```

1. In the **taskpane.js** file, replace `TODO8` with the following code to assign the event handlers for the **Add Slides** and four navigation buttons.

    ```js
    document.getElementById("add-slides").onclick = () => tryCatch(addSlides);
    document.getElementById("go-to-first-slide").onclick = () => clearMessage(goToFirstSlide);
    document.getElementById("go-to-next-slide").onclick = () => clearMessage(goToNextSlide);
    document.getElementById("go-to-previous-slide").onclick = () => clearMessage(goToPreviousSlide);
    document.getElementById("go-to-last-slide").onclick = () => clearMessage(goToLastSlide);
    ```

1. In the **taskpane.js** file, replace `TODO9` with the following code to define the `addSlides` and navigation functions. Each of these functions uses the `goToByIdAsync` method to select a slide based upon its position in the document (first, last, previous, and next).

    ```js
    async function addSlides() {
      await PowerPoint.run(async function (context) {
        context.presentation.slides.add();
        context.presentation.slides.add();

        await context.sync();

        goToLastSlide();
        setMessage("Success: Slides added.");
      });
    }

    function goToFirstSlide() {
      Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
        }
      });
    }

    function goToLastSlide() {
      Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
        }
      });
    }

    function goToPreviousSlide() {
      Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
        }
      });
    }

    function goToNextSlide() {
      Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          setMessage("Error: " + asyncResult.error.message);
        }
      });
    }
    ```

1. Save all your changes to the project.

### Test the add-in

1. Navigate to the root folder of the project.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. If the local web server isn't already running, complete the following steps to start the local web server and sideload your add-in.

    [!INCLUDE [alert use https](../includes/alert-use-https.md)]

    > [!TIP]
    > If you're testing your add-in on Mac, run the following command before proceeding. When you run this command, the local web server starts.
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    - To test your add-in in PowerPoint, run the following command in the root directory of your project. This starts the local web server (if it's not already running) and opens PowerPoint with your add-in loaded.

        ```command&nbsp;line
        npm start
        ```

    - To test your add-in in PowerPoint on a browser, run the following command in the root directory of your project. When you run this command, the local web server starts. Replace "{url}" with the URL of a PowerPoint document on your OneDrive or a SharePoint library to which you have permissions.

        [!INCLUDE [npm start:web command syntax](../includes/start-web-sideload-instructions.md)]

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button highlighted on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-yo-show-taskpane-button.png)

1. In the task pane, choose the **Add Slides** button. Two new slides are added to the document and the last slide in the document is selected and displayed.

    ![The Add Slides button highlighted in the add-in.](../images/powerpoint-tutorial-yo-add-slides.png)

1. In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.

    ![The Go to First Slide button highlighted in the add-in.](../images/powerpoint-tutorial-yo-go-to-first-slide.png)

1. In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.

    ![The Go to Next Slide button highlighted in the add-in.](../images/powerpoint-tutorial-yo-go-to-next-slide.png)

1. In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.

    ![The Go to Previous Slide button highlighted in the add-in.](../images/powerpoint-tutorial-yo-go-to-previous-slide.png)

1. In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.

    ![The Go to Last Slide button highlighted in the add-in.](../images/powerpoint-tutorial-yo-go-to-last-slide.png)

1. If the web server is running, run the following command when you want to stop the server.

    ```command&nbsp;line
    npm stop
    ```

## Code samples

- [Completed PowerPoint add-in tutorial](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/powerpoint-tutorial-yo): The result of completing this tutorial.

# [Visual Studio](#tab/visualstudio)

> [!TIP]
> If you want a completed version of this tutorial, head over to the [Office Add-ins samples repo on GitHub](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/powerpoint-tutorial).

## Prerequisites

- [Visual Studio 2019 or later](https://www.visualstudio.com/vs/), with the **Office/SharePoint development** workload installed.

    > [!NOTE]
    > If you've previously installed Visual Studio, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.

- Office connected to a Microsoft 365 subscription (including Office on the web).

    > [!NOTE]
    > If you don't already have Office, you can [join the Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program) to get a free, 90-day renewable Microsoft 365 subscription to use during development.

## Create your add-in project

Complete the following steps to create a PowerPoint add-in project using Visual Studio.

1. Choose **Create a new project**.

1. Using the search box, enter **add-in**. Choose **PowerPoint Web Add-in**, then select **Next**.

1. Name the project `HelloWorld`, and select **Create**.

1. In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.

1. Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.

     ![The Visual Studio Solution Explorer window showing HelloWorld and HelloWorldWeb, the two projects in the HelloWorld solution.](../images/powerpoint-tutorial-solution-explorer.png)

1. The following NuGet packages must be installed. Install them on the **HelloWorldWeb** project using the **NuGet Package Manager** in Visual Studio. See Visual Studio help for instructions. The second of these may be installed automatically when you install the first.

   - Microsoft.AspNet.WebApi.WebHost
   - Microsoft.AspNet.WebApi.Core

   > [!IMPORTANT]
   > When you're using the **NuGet Package Manager** to install these packages, do **not** install the recommended update to jQuery. The jQuery version installed with your Visual Studio solution matches the jQuery call within the solution files.

1. Use the **NuGet Package Manager** to update the Newtonsoft.Json package to version 13.0.3 or later. Then delete the **app.config** file if it was added to the **HelloWorld** project.

### Explore the Visual Studio solution

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### Update code

Edit the add-in code as follows to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.

1. **Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the add-slides and go-to-slide buttons. -->
        </div>
    </div>
    ```

1. Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.

    ```js
    (function () {
        "use strict";

        let messageBanner;

        Office.onReady(function () {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it.
                const element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for add-slides and the four navigation buttons.
            });
        });

        // TODO2: Define the insertImage function.

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the addSlides and navigation functions.

        // Helper function for displaying notifications.
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```

## Insert an image

Complete the following steps to add code that retrieves the [Bing](https://www.bing.com) photo of the day and inserts that image into a slide.

1. Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.

    ![The Visual Studio Solution Explorer window showing the Controllers folder highlighted in the HelloWorldWeb project.](../images/powerpoint-tutorial-solution-explorer-controllers.png)

1. Right-click the **Controllers** folder and select **Add** > **New Scaffolded Item...**.

1. In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.

1. In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button. Visual Studio creates and opens the **PhotoController.cs** file.

    > [!IMPORTANT]
    > The scaffolding process doesn't complete properly on some versions of Visual Studio after version 16.10.3. If you have the **Global.asax** and **./App_Start/WebApiConfig.cs** files, then skip to step 6.
    >
    > ![The Visual Studio Solution Explorer window showing the scaffolded files highlighted in the HelloWorldWeb project.](../images/powerpoint-tutorial-solution-explorer-scaffolded.png)

1. If you're missing scaffolding files from the **HelloWorldWeb** project, add them as follows.

    1. Using Solution Explorer, add a new folder named **App_Start** to the **HelloWorldWeb** project.

    1. Right-click the **App_Start** folder and select **Add** > **Class...**.

    1. In the **Add New Item** dialog, name the file **WebApiConfig.cs** then choose the **Add** button.

    1. Replace the entire contents of the **WebApiConfig.cs** file with the following code.

        ```cs
        using System;
        using System.Collections.Generic;
        using System.Linq;
        using System.Web;
        using System.Web.Http;
        
        namespace HelloWorldWeb.App_Start
        {
            public static class WebApiConfig
            {
                public static void Register(HttpConfiguration config)
                {
                    config.MapHttpAttributeRoutes();
        
                    config.Routes.MapHttpRoute(
                        name: "DefaultApi",
                        routeTemplate: "api/{controller}/{id}",
                        defaults: new { id = RouteParameter.Optional }
                    );
                }
            }
        }
        ```

    1. In the Solution Explorer, right-click the **HelloWorldWeb** project and select **Add** > **New Item...**.

    1. In the **Add New Item** dialog, search for "global", select **Global Application Class**, then choose the **Add** button. By default, the file is named **Global.asax**.

    1. Replace the entire contents of the **Global.asax.cs** file with the following code.

        ```cs
        using HelloWorldWeb.App_Start;
        using System;
        using System.Collections.Generic;
        using System.Linq;
        using System.Web;
        using System.Web.Http;
        using System.Web.Security;
        using System.Web.SessionState;
        
        namespace HelloWorldWeb
        {
            public class WebApiApplication : System.Web.HttpApplication
            {
                protected void Application_Start()
                {
                    GlobalConfiguration.Configure(WebApiConfig.Register);
                }
            }
        }
        ```

    1. In the Solution Explorer, right-click the **Global.asax** file and choose **View Markup**.

    1. Replace the entire contents of the **Global.asax** file with the following code.

        ```XML
        <%@ Application Codebehind="Global.asax.cs" Inherits="HelloWorldWeb.WebApiApplication" Language="C#" %>
        ```

1. Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64-encoded string. When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64-encoded string.

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the XML response and get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64-encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

1. In the **Home.html** file, replace `TODO1` with the following markup. This markup defines the **Insert Image** button that will appear within the add-in's task pane.

    ```html
    <button class="Button Button--primary" id="insert-image">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Insert Image</span>
        <span class="Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

1. In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.

    ```js
    $('#insert-image').click(insertImage);
    ```

1. In the **Home.js** file, replace `TODO2` with the following code to define the `insertImage` function. This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.

    ```js
    function insertImage() {
        // Get image from web service (as a Base64-encoded string).
        $.ajax({
            url: "/api/photo/",
            dataType: "text",
            success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

1. In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function. This function uses the Office JavaScript API to insert the image into the document. Note:

    - The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsync` request indicates the type of data being inserted.

    - The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### Test the add-in

1. Using Visual Studio, test the newly created PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button highlighted on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.

    ![The PowerPoint add-in with the Insert Image button highlighted.](../images/powerpoint-tutorial-insert-image-button.png)

    > [!NOTE]
    > If you get an error "Could not find file [...]\bin\roslyn\csc.exe", then do the following:
    >
    > 1. Open the **.\Web.config** file.
    > 1. Find the **\<compiler\>** node for the .cs `extension`, then remove the `type` attribute and its value.
    > 1. Save the file.

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Customize user interface (UI) elements

Complete the following steps to add markup that customizes the task pane UI.

1. In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane. Note:

    - The styles that begin with `ms-` are defined by [Fabric Core in Office Add-ins](../design/fabric-core.md), a JavaScript front-end framework for building user experiences for Office. The **Home.html** file includes a reference to the Fabric Core stylesheet.

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

1. In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.

### Test the add-in

1. Using Visual Studio, test the PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button highlighted on the PowerPoint Home ribbon.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. Notice that the task pane now contains a header section and title, and no longer contains a footer section.

    ![The PowerPoint add-in with Insert Image button.](../images/powerpoint-tutorial-new-task-pane-ui.png)

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Insert text

Complete the following steps to add code that inserts text into the title slide which contains the [Bing](https://www.bing.com) photo of the day.

1. In the **Home.html** file, replace `TODO3` with the following markup. This markup defines the **Insert Text** button that will appear within the add-in's task pane.

    ```html
        <br /><br />
        <button class="Button Button--primary" id="insert-text">
            <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="Button-label">Insert Text</span>
            <span class="Button-description">Inserts text into the slide.</span>
        </button>
    ```

1. In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.

    ```js
    $('#insert-text').click(insertText);
    ```

1. In the **Home.js** file, replace `TODO5` with the following code to define the `insertText` function. This function inserts text into the current slide.

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### Test the add-in

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.

    ![The selected PowerPoint title slide and the Insert Image button highlighted in the add-in.](../images/powerpoint-tutorial-insert-image-slide-design.png)

1. Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.

    ![The selected PowerPoint title slide with the Insert Text button highlighted in the add-in.](../images/powerpoint-tutorial-insert-text.png)

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Get slide metadata

Complete the following steps to add code that retrieves metadata for the selected slide.

1. In the **Home.html** file, replace `TODO4` with the following markup. This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.

    ```html
    <br /><br />
    <button class="Button Button--primary" id="get-slide-metadata">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get Slide Metadata</span>
        <span class="Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

1. In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

1. In the **Home.js** file, replace `TODO7` with the following code to define the `getSlideMetadata` function. This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

### Test the add-in

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button in Visual Studio.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button on the PowerPoint Home ribbon.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written to the popup dialog window at the bottom of the task pane. In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.

    ![The Get Slide Metadata button highlighted in the add-in.](../images/powerpoint-tutorial-get-slide-metadata.png)

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button in Visual Studio.](../images/powerpoint-tutorial-stop.png)

## Navigate between slides

Complete the following steps to add code that navigates between the slides of a document.

1. In the **Home.html** file, replace `TODO5` with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.

    ```html
    <br /><br />
    <button class="Button Button--primary" id="add-slides">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Add Slides</span>
        <span class="Button-description">Adds 2 slides.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-first-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to First Slide</span>
        <span class="Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-next-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Next Slide</span>
        <span class="Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-previous-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Previous Slide</span>
        <span class="Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-last-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Last Slide</span>
        <span class="Button-description">Go to the last slide.</span>
    </button>
    ```

1. In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the **Add Slides** and four navigation buttons.

    ```js
    $('#add-slides').click(addSlides);
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

1. In the **Home.js** file, replace `TODO9` with the following code to define the `addSlides` and navigation functions. Each of these functions uses the `goToByIdAsync` method to select a slide based upon its position in the document (first, last, previous, and next).

    ```js
    async function addSlides() {
        await PowerPoint.run(async function (context) {
            context.presentation.slides.add();
            context.presentation.slides.add();

            await context.sync();

            showNotification("Success", "Slides added.");
            goToLastSlide();
        });
    }

    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### Test the add-in

1. Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed on the ribbon. The add-in will be hosted locally on IIS.

    ![The Start button highlighted on the Visual Studio toolbar.](../images/powerpoint-tutorial-start.png)

1. If the add-in task pane isn't already open in PowerPoint, select the **Show Taskpane** button on the ribbon to open it.

    ![The Show Taskpane button highlighted on the Home ribbon in PowerPoint.](../images/powerpoint-tutorial-show-taskpane-button.png)

1. In the task pane, choose the **Add Slides** button. Two new slides are added to the document and the last slide in the document is selected and displayed.

    ![The Add Slides button highlighted in the add-in.](../images/powerpoint-tutorial-add-slides-1.png)

1. In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.

    ![The Go to First Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-first-slide.png)

1. In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.

    ![The Go to Next Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-next-slide.png)

1. In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.

    ![The Go to Previous Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-previous-slide.png)

1. In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.

    ![The Go to Last Slide button highlighted in the add-in.](../images/powerpoint-tutorial-go-to-last-slide.png)

1. In Visual Studio, stop the add-in by pressing **Shift+F5** or choosing the **Stop** button. PowerPoint will automatically close when the add-in is stopped.

    ![The Stop button highlighted on the Visual Studio toolbar.](../images/powerpoint-tutorial-stop.png)

## Code samples

- [Completed PowerPoint add-in tutorial](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/tutorials/powerpoint-tutorial): The result of completing this tutorial.

---

## Next steps

In this tutorial, you've created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides. To learn more about building PowerPoint add-ins, continue to the following article.

> [!div class="nextstepaction"]
> [PowerPoint add-ins overview](../powerpoint/powerpoint-add-ins.md)

## See also

- [Office Add-ins platform overview](../overview/office-add-ins.md)
- [Develop Office Add-ins](../develop/develop-overview.md)
