/* Console History by RL Vision *
 * https://github.com/rlv-dan   */
(function () {

	if (console && console.log) {	// Make sure the console object is available
		
		console.historyData = [];
		
		// Restore previous session (if there is one)
		// The sessionStorage object is a convenient choice because the browser automatically clears it for us.
		var currentSessionStorage = [];
		if(sessionStorage && sessionStorage.consoleHistory) {
			currentSessionStorage = JSON.parse(sessionStorage.consoleHistory);
		}

		// Keep a reference to the original console functions that we intend to overwrite
		window.originalConsole = {
			// Assigning a native function to a var could throw TypeError, thus we need to use bind
			log: console.log.bind(console),
			info: console.info.bind(console),
			warn: console.warn.bind(console),
			error: console.error.bind(console),
		}

		// Function to save console calls in our history data objects
		var makeHistory = function(type, args) {
			args= Array.prototype.slice.call(args);
			console.historyData.push( { "type": type, "output": args, "timestamp": new Date() } );
			if(sessionStorage) {
				currentSessionStorage.push( { "type": type, "output": args, "timestamp": new Date() } );
				sessionStorage.consoleHistory = JSON.stringify(currentSessionStorage);
			}
		}

		// Override console functions with out own implementations
		// These will first record the call arguments and then pass the call on to the real console object
		console.log = function() {
			makeHistory("log", arguments);
			return originalConsole.log.apply(this, arguments);
		}
		console.info = function() {
			makeHistory("info", arguments);
			return originalConsole.info.apply(this, arguments);
		}
		console.warn = function() {
			makeHistory("warn", arguments);
			return originalConsole.warn.apply(this, arguments);
		}	
		console.error = function() {
			makeHistory("error", arguments);
			return originalConsole.error.apply(this, arguments);
		}

		// Recall history by calling console.history()
		// Argument entireSession determines if recalling since beginning of current page load or since beginning of the current session
		console.history = function( entireSession ) {
		
			// Select data source
			var source = console.historyData;
			if(entireSession && sessionStorage && sessionStorage.consoleHistory) {
				source = JSON.parse(sessionStorage.consoleHistory);
			}
			
			// Print history
			for(var n=0; n<source.length; n++) {
				switch(source[n].type) {
					case "log":
						originalConsole.log.apply(null, source[n].output);
						break;
					case "info":
						originalConsole.info.apply(null, source[n].output);
						break;
					case "warn":
						originalConsole.warn.apply(null, source[n].output);
						break;
					case "error":
						originalConsole.error.apply(null, source[n].output);
						break;
				}
			}
		}
	
	}
})(); // Initialize using an Immediately-Invoked Function Expression (IIFE)
