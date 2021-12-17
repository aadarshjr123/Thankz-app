import { CardFactory } from "botbuilder";

const Takeatour = CardFactory.adaptiveCard({
  version: "1.3",

  type: "AdaptiveCard",
  body: [
    {
      type: "TextBlock",
      size: "Medium",
      weight: "Bolder",
      text: "About",
    },
    {
      type: "ColumnSet",
      columns: [
        {
          type: "Column",
          items: [
            {
              type: "TextBlock",
              weight: "Bolder",
              text: "Award List",
              wrap: true,
            },
          ],
          width: "stretch",
        },
      ],
    },
    {
      type: "TextBlock",
      text: "The award list tab helps us to view tha awards received and sent ",
      wrap: true,
    },
    {
      type: "TextBlock",
      text: "Leader Board",
      wrap: true,
      weight: "Bolder",
    },
    {
      type: "TextBlock",
      text:
        "The leader Board tab helps us to view the most appreciated users",
      wrap: true,
    },
  ],
  $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
});
export default Takeatour;
