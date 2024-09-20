/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // document.getElementById("sideload-msg").style.display = "none";
    // document.getElementById("app-body").style.display = "flex";

    const dummyCardsData = [
      {
        imgSrc: "../../assets/logo-filled.png",
        senderName: "Abhinav Gandotra",
        emailTitle: "Meeting regarding finalization of work demo.",
        emailDescription: "Here we will give a quick summary of what have been done! This will help the reader a glimpse about the thread.",
      },
      {
        imgSrc: "../../assets/logo-filled.png",
        senderName: "Jane Doe",
        emailTitle: "Meeting regarding finalization of work demo.",
        emailDescription: "Here we will give a quick summary of what have been done! This will help the reader a glimpse about the thread.",
      },
      {
        imgSrc: "../../assets/logo-filled.png",
        senderName: "John Doe",
        emailTitle: "Meeting regarding finalization of work demo.",
        emailDescription: "Here we will give a quick summary of what have been done! This will help the reader a glimpse about the thread.",
      },
      {
        imgSrc: "../../assets/logo-filled.png",
        senderName: "Jane Doe",
        emailTitle: "Meeting regarding finalization of work demo.",
        emailDescription: "Here we will give a quick summary of what have been done! This will help the reader a glimpse about the thread.",
      }
    ]


    let cardsHtml = "";
    dummyCardsData.forEach((card) => {
      cardsHtml += `<fluent-card class="card">
                <div class="card-icon">
                    <img src="https://outlook.office365.com/10bb6ab6-495b-4322-890d-71c846ad427e" alt="Email Icon" />
                </div>
                <div class="card-content">
                    <fluent-card-header>
                        <fluent-card-header-title class="sender-name">${card.senderName}</fluent-card-header-title>
                    </fluent-card-header>
                    <fluent-card-body>
                        <p class="email-title">Meeting regarding finalization of work demo.</p>
                        <p class="email-description">Here we will give a quick summary of what have been done! This will help the reader a glimpse about the thread.</p>
                        <fluent-badge appearance="accent">New</fluent-badge>
                    </fluent-card-body>
                </div>
            </fluent-card>`;
    });

    const cards = document.getElementById("cards");
    cards.innerHTML = cardsHtml;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */
  // Here we can write any logic if we want
}
