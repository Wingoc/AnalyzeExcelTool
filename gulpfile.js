var gulp = require('gulp');

// 组件
var jshint = require('gulp-jshint');
var uglify = require('gulp-uglify');
var htmlmini = require('gulp-minify-html');
var minify = require('gulp-minify-css');

gulp.task('default', function(){
	console.log("This is my first Project!");
});


// 检查脚本js
gulp.task('checkjs', function(){
	gulp.src('./js/main.js')
		.pipe(jshint())
		.pipe(jshint.reporter('default'));
});

// 压缩脚本js
gulp.task('uglifyjs', function(){
	gulp.src('./js/main.js')   // 获取js文件，过滤.min.js文件
		.pipe(uglify())
		.pipe(gulp.dest('./dest/js/'));
});

// 压缩html文件
gulp.task('htmlmini', function(){
	gulp.src('./index.html')
		.pipe(htmlmini())
		.pipe(gulp.dest('./dest/'));
});

// 压缩css文件
gulp.task('cssmini', function(){
	gulp.src('./css/main.css')
		.pipe(minify())
		.pipe(gulp.dest('./dest/css/'));
});

// 监视文件变化
// gulp.task('mywatch', function () {
// 	gulp.watch('./js/main.js', ['checkjs', 'uglifyjs']);
// });
gulp.task('mywatch2', function(){
	gulp.watch('./css/main.css', ['cssmini']);
});
