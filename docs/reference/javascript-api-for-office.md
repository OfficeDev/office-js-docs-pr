---
layout: LandingPage
ms.topic: landing-page
title: Office JavaScript API reference documentation
description: Learn about the Office JavaScript APIs.
ms.date: 10/14/2020
ms.localizationpriority: high
---

# API reference documentation

An add-in can use the Office JavaScript APIs to interact with objects in Office client applications. 

<ul>
    <li><b>Application-specific</b> APIs provide strongly-typed objects that can be used to interact with objects that are native to a specific Office application.</li>
    <li><b>Common</b> APIs can be used to access features such as UI, dialogs, and client settings that are common across multiple types of Office applications.</li>
</ul>

You should use application-specific APIs whenever feasible, and use Common APIs only for scenarios that aren't supported by application-specific APIs. For more detailed information about these two API models, see <a href="../develop/develop-overview.md#api-models">Develop Office Add-ins</a>.

<h2>API reference</h2>

<ul class="panelContent cardsF cols cols3">
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/excel"><img src="../images/index/logo-excel.svg" alt="Excel API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Excel API reference</h3>
                        <p><a href="/javascript/api/excel">JavaScript APIs for building Excel add-ins.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Outlook API reference</h3>
                        <p><a href="/javascript/api/outlook">JavaScript APIs for building Outlook add-ins.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/word"><img src="../images/index/logo-word.svg" alt="Word API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Word API reference</h3>
                        <p><a href="/javascript/api/word">JavaScript APIs for building Word add-ins.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>PowerPoint API reference</h3>
                        <p><a href="/javascript/api/powerpoint">JavaScript APIs for building PowerPoint add-ins.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>OneNote API reference</h3>
                        <p><a href="/javascript/api/onenote">JavaScript APIs for building OneNote add-ins.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/office"><img src="../images/index-landing-page/i_code-blocks.svg" alt="reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Common API reference</h3>
                        <p><a href="/javascript/api/office">JavaScript APIs that can be used by any Office Add-in.</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
</ul>

<b>Note</b>: There's currently no application-specific JavaScript API for Project; you'll use Common APIs to create Project add-ins. Additionally, the application-specific API for PowerPoint is very limited in scope; you'll mainly use Common APIs to create PowerPoint add-ins.
