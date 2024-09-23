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

    const cardElements = document.getElementsByClassName("card");
    for (const element of cardElements)
    {
      element.onclick = showAnalysis;
    }

    // Rendering the chart
    (async () => {
      Highcharts.chart('container', {
        chart: {
          type: 'area'
        },
        title: {
          text: null
        },
        xAxis: {
          type: "datetime",
          title: {
            text: null
          },
          categories: [
            '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', 
            '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', 
            '21', '22', '23', '24', '25', '26', '27'
          ]
        },
        yAxis: {
          title: {
            text: null
          },
          min: -1,
          max: 1
        },
        series: [{
          name: 'Overall Tone Score',
          data: [
            0.5, -0.2, 0.7, -0.1, 0.3, -0.4, 0.8, -0.6, 0.9, -0.3,
            1.0, -0.5, -0.7, -0.8, -0.8, 0.6, -0.9, 0.7, -0.7, 0.9, -1.0,
            0.3, -0.6, 0.4, -0.8, 0.5, -0.3, -0.3, 0.7
          ],
          color: '#28a745',
          negativeColor: '#dc3545'
        }],
        tooltip: {
          xDateFormat: '%A, %b %e, %Y', // Format the date in the tooltip
          pointFormat: '{point.x:%b %e, %Y}: ({point.y:.2f})' // Custom format for the tooltip
        },
        plotOptions: {
          area: {
            marker: {
              radius: 2
            },
          },
          column: {
            borderColor: 'transparent',
            colorByPoint: false,
            zones: [{
              value: 0, // Anything below zero
              color: '#dc3545' // Red for negative values
            }, {
              value: 1, // Anything equal to or above zero
              color: '#28a745' // Green for positive values
            }]
          }
        }
      });
    })();
  }

  document.getElementById("back-btn").onclick = GoBack;
});

export async function showAnalysis() {
  /**
   * Insert your Outlook code here
   */
  // Here we can write any logic if we want
  document.getElementById("priority-list-container").classList.add("hidden");
  document.getElementById("sentiment-container").classList.remove("hidden");
}

export async function GoBack() {
  /**
   * Insert your Outlook code here
   */
  // Here we can write any logic if we want
  document.getElementById("priority-list-container").classList.remove("hidden");
  document.getElementById("sentiment-container").classList.add("hidden");
}
