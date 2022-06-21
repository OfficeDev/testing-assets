Office.actions.associate('taskpane.openTaskpane', function () {
	return Office.addin.showAsTaskpane()
		.then(function () {
			return;
		})
		.catch(function (error) {
			return error.code;
		});
});

Office.actions.associate('taskpane.updateSharedValue', function () {
	g_sharedAppData.value = 2021;
});

Office.actions.associate('HIDETASKPANE', function () {
	return Office.addin.hide()
		.then(function () {
			return;
		})
		.catch(function (error) {
			return error.code;
		});
});

function SetRuntimeVisibleHelper(visible) {
	var p;
	if (visible) {
		p = Office.addin.showAsTaskpane();
	}
	else {
		p = Office.addin.hide();
	}

	return p.then(function () {
		return visible;
	})
		.catch(function (error) {
			return error.code;
		});
}

function SetStartupBehaviorHelper(state) {
	return Office.addin.setStartupBehavior(state)
		.then(function () {
			return state;
		})
		.catch(function (error) {
			return error.code;
		});
}