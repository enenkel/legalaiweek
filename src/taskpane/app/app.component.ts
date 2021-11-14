import { Component } from "@angular/core";

/* global Word */

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  welcomeMessage = "Team Oppenländer";

  async readText() {
    return Word.run(async (context) => {
      console.log("BODY -> ", context.document.body);
    });
  }

  async run() {
    return Word.run(async (context) => {
      /**
       * Insert your Word code here
       */

      const text = `An das\v□	Amtsgericht\v□	Landgericht\vin _________________________\v\v` + 
      `In dem Rechtsstreit \v Kläger ./. Beklagter\v\vAz: _________________________\v\vzeige ich an, dass der Beklagte vom Unterzeichner vertreten wird. Namens und in Vollmacht des Beklagten wird mitgeteilt, dass dieser sich gegen die Klage verteidigen will.\v\vNamens und in Vollmacht des Beklagten werde ich in der mündlichen Verhandlung beantragen,\v\v \v   die Klage abzuweisen.\vZur Klageerwiderung wird wie folgt vorgetragen:\v\vDie Klage ist\v\v□	bereits unzulässig, ungeachtet dessen aber auch unbegründet.\v□	zwar zulässig, nicht jedoch begründet.\vDem Kläger steht der geltend gemachte Anspruch nicht zu, weil\v\v□	der von ihm dargestellte Sachverhalt nicht dem tatsächlichen Geschehen entspricht.\v□	der von dem Kläger dargestellte Sachverhalt teilweise nicht dem tatsächlichen Geschehen entspricht und unter Berücksichtigung des tatsächlichen Geschehensablaufes der geltend gemachte Anspruch nicht begründet werden kann.\v□	der vom Kläger geltend gemachte Sachverhalt zwar den Tatsachen entspricht, der geltend gemachte Anspruch hieraus jedoch nicht hergeleitet werden kann.\v□	der mit der Klage geltend gemachte Anspruch zwar ursprünglich bestanden hat, jedoch jetzt nicht mehr besteht, weil _________________________\v□	_________________________\vIm Einzelnen ist hierzu Folgendes vorzutragen:`;

      // insert a paragraph at the end of the document.
      const paragraph = context.document.body.insertParagraph(text, Word.InsertLocation.end);

      const response = await fetch('https://reqres.in/api/users/2')
      .then(response => response.json());
      const paragraph2 = context.document.body.insertParagraph("Response: " + JSON.stringify(response), Word.InsertLocation.end);

      await context.sync();
    });
  }
}
