(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
	}

	$(document).ready(function () {
		app.initialize();

		RichApiTest.buildUI(
			document.getElementById("DivTests"),
			"PowerPointTests",
			[
				"/Scripts/EditorIntelliSense/PowerPointTests.txt",
				"/Scripts/EditorIntelliSense/PowerPoint.txt",
				"/Scripts/EditorIntelliSense/Office.Runtime.txt",
				"/Scripts/EditorIntelliSense/Office.Core.txt",
				"/Scripts/EditorIntelliSense/RichApiTest.Core.txt",
				"/Scripts/EditorIntelliSense/Helpers.txt",
				"/Scripts/EditorIntelliSense/jquery.txt",
				"/Scripts/EditorIntelliSense/office-js.txt",
				"/Scripts/EditorIntelliSense/Office.Core.Test.txt",
			],
			{
				"PowerPointTests": "/Scripts/EditorIntelliSense/PowerPointTests.sources.txt",
				"OfficeCoreTest": "/Scripts/EditorIntelliSense/Office.Core.Test.sources.txt"
			}
		);

		RichApiTest.appendTests("OfficeCoreTest");
	});

})();

