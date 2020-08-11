CustomFunctions.associate('GetValue', function() {
	if (typeof(g_sharedAppData) === 'object') {
		return g_sharedAppData.value;
	}

	return null;
});

CustomFunctions.associate('SetValue', function(value) {
	if (typeof(g_sharedAppData) === 'object') {
		g_sharedAppData.value = value;
		return value;
	}

	return null;
});
