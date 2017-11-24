const math = require('mathjs');

class Progress {
	constructor(maximum) {
		this.max = maximum;
		this.progress = 0;
	}

	tick(update) {
		process.stdout.write('\r\x1b[K');
		if (typeof update !== 'undefined') {
			this.progress += update;
			process.stdout.write(this.max + '/' + math.floor((this.progress / this.max * 100) * 100) / 100);
		}
	}
}

module.exports = Progress;