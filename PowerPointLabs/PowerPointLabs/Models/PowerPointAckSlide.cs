using System;
using System.IO;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointAckSlide : PowerPointSlide
    {
#pragma warning disable 0618
        private const string PptLabsAckSlideName = "PPTLabsAcknowledgementSlide";

        //public const string teststr = "[  {    \"_id\": \"55700fe523d780a378b3d3c0\",    \"index\": 0,    \"guid\": \"4b67bc88-3826-4b1c-8912-fc0b47b5f3c7\",    \"isActive\": false,    \"balance\": \"$3,604.58\",    \"picture\": \"http://placehold.it/32x32\",    \"age\": 25,    \"eyeColor\": \"brown\",    \"name\": \"Mendoza Cote\",    \"gender\": \"male\",    \"company\": \"AQUAZURE\",    \"email\": \"mendozacote@aquazure.com\",    \"phone\": \"+1 (959) 566-2966\",    \"address\": \"659 Debevoise Avenue, Eggertsville, Maine, 7356\",    \"about\": \"Irure exercitation exercitation ut culpa do officia anim adipisicing mollit. Eu elit reprehenderit proident fugiat. Exercitation id labore in sunt aute ea dolore in officia aliqua do amet adipisicing. Enim nisi elit occaecat esse ipsum ad reprehenderit ea laborum exercitation enim nostrud reprehenderit. Minim velit id enim aliqua nostrud amet nisi est minim amet duis mollit pariatur dolore. Officia mollit enim qui ea. Laborum incididunt minim do deserunt nulla laboris proident esse pariatur enim ipsum dolore id deserunt.\\\",    \"registered\": \"2014-08-06T19:40:15 -08:00\",    \"latitude\": 30.318416,    \"longitude\": 124.322333,    \"tags\": [      \"laborum\",      \"fugiat\",      \"reprehenderit\",      \"cillum\",      \"incididunt\",      \"duis\",      \"ipsum\"    ],    \"friends\": [      {        \"id\": 0,        \"name\": \"Imogene Hays\"      },      {        \"id\": 1,        \"name\": \"Day Durham\"      },      {        \"id\": 2,        \"name\": \"Myrtle Curtis\"      }    ],    \"greeting\": \"Hello, Mendoza Cote! You have 3 unread messages.\",    \"favoriteFruit\": \"apple\"  },  {    \"_id\": \"55700fe58bce172ee6e5c272\",    \"index\": 1,    \"guid\": \"4d5b131a-0a95-40e0-a85b-e84d1c7b0f6e\",    \"isActive\": false,    \"balance\": \"$2,468.79\",    \"picture\": \"http://placehold.it/32x32\",    \"age\": 29,    \"eyeColor\": \"blue\",    \"name\": \"Franks Boyd\",    \"gender\": \"male\",    \"company\": \"SONIQUE\",    \"email\": \"franksboyd@sonique.com\",    \"phone\": \"+1 (923) 536-3871\",    \"address\": \"387 Hendrickson Place, Fontanelle, Nebraska, 4922\",    \"about\": \"Laborum laboris Lorem aute nostrud dolore. Nulla amet exercitation esse cillum nostrud incididunt nostrud adipisicing reprehenderit esse ex qui veniam. Culpa culpa irure dolor anim. Ad ipsum pariatur proident quis irure esse amet qui sint nostrud mollit culpa eiusmod occaecat.\\\",    \"registered\": \"2014-07-19T01:40:53 -08:00\",    \"latitude\": -10.395386,    \"longitude\": 104.923394,    \"tags\": [      \"quis\",      \"irure\",      \"aliqua\",      \"ullamco\",      \"voluptate\",      \"dolor\",      \"quis\"    ],    \"friends\": [      {        \"id\": 0,        \"name\": \"Marta Baird\"      },      {        \"id\": 1,        \"name\": \"Cleo Carter\"      },      {        \"id\": 2,        \"name\": \"Gray Yates\"      }    ],    \"greeting\": \"Hello, Franks Boyd! You have 2 unread messages.\",    \"favoriteFruit\": \"apple\"  },  {    \"_id\": \"55700fe5366e91b52fd34b96\",    \"index\": 2,    \"guid\": \"8fb0e670-62f9-4758-9ee1-e621b62f7a1a\",    \"isActive\": false,    \"balance\": \"$3,477.39\",    \"picture\": \"http://placehold.it/32x32\",    \"age\": 22,    \"eyeColor\": \"green\",    \"name\": \"Amparo Nieves\",    \"gender\": \"female\",    \"company\": \"DATAGEN\",    \"email\": \"amparonieves@datagen.com\",    \"phone\": \"+1 (864) 594-2274\",    \"address\": \"332 Tabor Court, Hemlock, Illinois, 5286\",    \"about\": \"Ipsum aliqua nulla voluptate aute commodo do laboris reprehenderit eu labore veniam dolor. Eu cupidatat adipisicing do labore dolor magna sit velit nisi voluptate non et adipisicing. Occaecat culpa aliquip do amet nulla veniam sint irure velit magna mollit consectetur veniam reprehenderit. Labore sit excepteur incididunt ex sunt aliquip Lorem aute incididunt.\\\",    \"registered\": \"2014-09-04T23:35:22 -08:00\",    \"latitude\": 23.184812,    \"longitude\": -41.53521,    \"tags\": [      \"occaecat\",      \"enim\",      \"labore\",      \"sint\",      \"consectetur\",      \"cillum\",      \"velit\"    ],    \"friends\": [      {        \"id\": 0,        \"name\": \"Lydia Howe\"      },      {        \"id\": 1,        \"name\": \"Rosales Johnson\"      },      {        \"id\": 2,        \"name\": \"Mcknight Brennan\"      }    ],    \"greeting\": \"Hello, Amparo Nieves! You have 1 unread messages.\",    \"favoriteFruit\": \"banana\"  },  {    \"_id\": \"55700fe54e2b472964afba65\",    \"index\": 3,    \"guid\": \"ebce1253-2ac6-42ee-af0d-b9f2f79d5dbf\",    \"isActive\": true,    \"balance\": \"$2,630.07\",    \"picture\": \"http://placehold.it/32x32\",    \"age\": 37,    \"eyeColor\": \"brown\",    \"name\": \"Liza Mercer\",    \"gender\": \"female\",    \"company\": \"PHORMULA\",    \"email\": \"lizamercer@phormula.com\",    \"phone\": \"+1 (857) 409-3505\",    \"address\": \"464 Ludlam Place, Galesville, Alaska, 5665\",    \"about\": \"Irure ad elit exercitation dolore ipsum sit exercitation officia dolor. Ut ea magna consectetur id do dolore sunt tempor occaecat non. Pariatur commodo eu duis commodo voluptate commodo reprehenderit ex magna. Mollit minim qui pariatur officia deserunt labore irure consequat et consectetur dolor laboris.\\\",    \"registered\": \"2015-05-07T09:08:54 -08:00\",    \"latitude\": 40.452527,    \"longitude\": -119.491377,    \"tags\": [      \"laborum\",      \"proident\",      \"aute\",      \"nisi\",      \"aliqua\",      \"pariatur\",      \"pariatur\"    ],    \"friends\": [      {        \"id\": 0,        \"name\": \"Church Mosley\"      },      {        \"id\": 1,        \"name\": \"Simmons Duran\"      },      {        \"id\": 2,        \"name\": \"Hull Dillon\"      }    ],    \"greeting\": \"Hello, Liza Mercer! You have 7 unread messages.\",    \"favoriteFruit\": \"strawberry\"  },  {    \"_id\": \"55700fe5498431918b33ad03\",    \"index\": 4,    \"guid\": \"db038e69-8754-4cac-918e-a3594f408d3b\",    \"isActive\": true,    \"balance\": \"$3,944.73\",    \"picture\": \"http://placehold.it/32x32\",    \"age\": 38,    \"eyeColor\": \"brown\",    \"name\": \"Jolene Mcconnell\",    \"gender\": \"female\",    \"company\": \"AQUASURE\",    \"email\": \"jolenemcconnell@aquasure.com\",    \"phone\": \"+1 (887) 572-3136\",    \"address\": \"213 Ira Court, Nash, New Jersey, 5231\",    \"about\": \"Et proident id magna Lorem aliqua velit exercitation. Aliquip tempor tempor labore voluptate adipisicing. Commodo dolore cupidatat adipisicing aute ea aute laboris commodo nulla sunt labore est. Eu eu ullamco quis commodo tempor aliquip. In nisi et deserunt eiusmod nostrud incididunt. Consequat nostrud nulla cupidatat pariatur laboris adipisicing in sint dolor laboris.\\\",    \"registered\": \"2014-11-17T23:25:29 -08:00\",    \"latitude\": 45.885005,    \"longitude\": -87.388396,    \"tags\": [      \"et\",      \"laboris\",      \"tempor\",      \"ex\",      \"laborum\",      \"in\",      \"id\"    ],    \"friends\": [      {        \"id\": 0,        \"name\": \"Moses Calhoun\"      },      {        \"id\": 1,        \"name\": \"Judy Berg\"      },      {        \"id\": 2,        \"name\": \"Maureen Washington\"      }    ],    \"greeting\": \"Hello, Jolene Mcconnell! You have 1 unread messages.\",    \"favoriteFruit\": \"apple\"  },  {    \"_id\": \"55700fe5703d684880dd0118\",    \"index\": 5,    \"guid\": \"e6af6119-b912-4473-add4-409acbea5e55\",    \"isActive\": true,    \"balance\": \"$2,811.17\",    \"picture\": \"http://placehold.it/32x32\",    \"age\": 24,    \"eyeColor\": \"brown\",    \"name\": \"Dickson Oneal\",    \"gender\": \"male\",    \"company\": \"UNI\",    \"email\": \"dicksononeal@uni.com\",    \"phone\": \"+1 (856) 578-2237\",    \"address\": \"838 Kane Street, Crown, Michigan, 9966\",    \"about\": \"Mollit consequat aute non adipisicing esse laboris voluptate ut ea ipsum. Ea do irure sunt ullamco ut do ullamco elit nulla esse cillum incididunt. Est esse adipisicing incididunt fugiat ex culpa duis labore sunt voluptate. Tempor duis aliquip dolor ad dolor dolor. Eu consectetur ut commodo culpa nisi cupidatat nisi in. Et exercitation laboris duis voluptate.\\\",    \"registered\": \"2015-05-06T07:27:11 -08:00\",    \"latitude\": 34.945217,    \"longitude\": 21.599256,    \"tags\": [      \"ut\",      \"labore\",      \"laboris\",      \"et\",      \"aliquip\",      \"id\",      \"adipisicing\"    ],    \"friends\": [      {        \"id\": 0,        \"name\": \"Letitia Bender\"      },      {        \"id\": 1,        \"name\": \"Casandra Mayer\"      },      {        \"id\": 2,        \"name\": \"Latisha Williamson\"      }    ],    \"greeting\": \"Hello, Dickson Oneal! You have 8 unread messages.\",    \"favoriteFruit\": \"banana\"  },  {    \"_id\": \"55700fe5a4aa9d5a14bda498\",    \"index\": 6,    \"guid\": \"872c495f-660a-42ff-96d6-d3eccfc574bd\",    \"isActive\": true,    \"balance\": \"$2,525.07\",    \"picture\": \"http://placehold.it/32x32\",    \"age\": 34,    \"eyeColor\": \"brown\",    \"name\": \"Delia Gillespie\",    \"gender\": \"female\",    \"company\": \"VORATAK\",    \"email\": \"deliagillespie@voratak.com\",    \"phone\": \"+1 (865) 549-2256\",    \"address\": \"731 Barlow Drive, Verdi, Colorado, 562\",    \"about\": \"Ex sunt adipisicing esse cillum ipsum consequat ipsum cillum. Minim do magna labore irure fugiat. Ullamco veniam dolor fugiat aute consequat cupidatat anim proident est id non aute labore qui. Est cillum commodo et ut aliquip et velit aliquip commodo sint sit exercitation reprehenderit dolore. Nulla eu reprehenderit ad qui quis non eiusmod.\\\",    \"registered\": \"2014-11-19T08:53:50 -08:00\",    \"latitude\": -71.62089,    \"longitude\": 169.773498,    \"tags\": [      \"officia\",      \"officia\",      \"occaecat\",      \"voluptate\",      \"adipisicing\",      \"nostrud\",      \"fugiat\"    ],    \"friends\": [      {        \"id\": 0,        \"name\": \"Nora Blankenship\"      },      {        \"id\": 1,        \"name\": \"Tammie Watson\"      },      {        \"id\": 2,        \"name\": \"Marlene Orr\"      }    ],    \"greeting\": \"Hello, Delia Gillespie! You have 2 unread messages.\",    \"favoriteFruit\": \"strawberry\"  }]";

        private PowerPointAckSlide(PowerPoint.Slide slide) : base(slide)
        {
            if (!IsAckSlide(slide.Name))
            {
                _slide.Name = PptLabsAckSlideName;
                String tempFileName = Path.GetTempFileName();
                Properties.Resources.Acknowledgement.Save(tempFileName);
                PowerPoint.Shape ackShape = _slide.Shapes.AddPicture(tempFileName, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 0, 0);
                _slide.SlideShowTransition.Hidden = Office.MsoTriState.msoTrue;

                ackShape.Left = (PowerPointPresentation.Current.SlideWidth - ackShape.Width) / 2;
                ackShape.Top = (PowerPointPresentation.Current.SlideHeight - ackShape.Height) / 2;

                //_slide.NotesPage.Shapes
                /*NotesPageText = teststr;
                Debug.WriteLine(teststr.Length);
                var output = NotesPageText;
                Debug.WriteLine(output.Length);
                //Debug.WriteLine("|" + output + "|");
                Debug.WriteLine(output == teststr);*/
            }
        }

        public static PowerPointSlide FromSlideFactory(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return null;
            }

            return new PowerPointAckSlide(slide);
        }

        public static bool IsAckSlide(PowerPointSlide slide)
        {
            if (slide == null)
            {
                return false;
            }

            return IsAckSlide(slide.Name);
        }

        public static bool IsAckSlide(PowerPoint.Slide slide)
        {
            if (slide == null)
            {
                return false;
            }

            return IsAckSlide(slide.Name);
        }

        private static bool IsAckSlide(string slideName)
        {
            return slideName == PptLabsAckSlideName;
        }
    }
}
