/* eslint-disable prettier/prettier */
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
        imgSrc: "../../assets/abhinav_img.jpeg",
        senderName: "Abhinav Gandotra",
        emailTitle: "Meeting regarding finalization of work demo.",
        summary: "Here we will give a quick summary of what have been done! This will help the reader a glimpse about the thread.",
        badges: [
          {
            badgeText: "Urgent",
            badgeColor: "ghost"
          },
          {
            badgeText: "High revenue",
            badgeColor: "accent"
          }
        ],
      },
      {
        imgSrc: "../../assets/zein_img.jpeg",
        senderName: "Rishabh Kalra",
        emailTitle: "Meeting regarding finalization of work demo.",
        summary: "Here we will give a quick summary of what have been done! This will help the reader a glimpse about the thread.",
        badges: [
          {
            badgeText: "Urgent",
            badgeColor: "red"
          },
          {
            badgeText: "High revenue",
            badgeColor: "blue"
          }
        ],
      },
      {
        imgSrc: "../../assets/priyam_img.jpeg",
        senderName: "Priyam Seth",
        emailTitle: "Need final quote of the project",
        summary: "The work has been finalized, just discussions left for the final quote.",
        badges: [
          {
            badgeText: "Urgent",
            badgeColor: "red"
          },
          {
            badgeText: "High revenue",
            badgeColor: "blue"
          }
        ],
      },
      {
        imgSrc: "../../assets/gaurav_img.jpeg",
        senderName: "Gaurav Sareen",
        emailTitle: "GXP Town Hall",
        summary: "Town Hall meeting - discussing the future of Products, new Customers and Ideas.",
        badges: [
          {
            badgeText: "Urgent",
            badgeColor: "red"
          },
          {
            badgeText: "High revenue",
            badgeColor: "blue"
          }
        ],
      }
    ];

    let cardsHtml = "";
    dummyCardsData.forEach((card) => {
      cardsHtml += `<fluent-card class="card">
                <div class="card-icon-container">
                    <img class="card-icon" src="${card.imgSrc}" alt="logo" title="Add-in logo" />
                </div>
                <div class="card-content">
                    <fluent-card-header>
                        <fluent-card-header-title class="sender-name">${card.senderName}</fluent-card-header-title>
                    </fluent-card-header>
                    <fluent-card-body>
                        <p class="email-title">${card.emailTitle}</p>
                        <p class="email-description">${card.summary}</p>
                        <div class="badges">
                        ${card.badges.map((badge) => `<fluent-badge appearance="${badge.badgeColor}">${badge.badgeText}</fluent-badge>`).join("")}
                        </div>
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
