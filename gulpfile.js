var gulp = require('gulp');
var cleanCompiledTypeScript = require('gulp-clean-compiled-typescript');
var del = require('del');
var sequence = require('run-sequence');

gulp.task('clean-jsmap', function () {
    return gulp.src('./**/*.ts')
        .pipe(cleanCompiledTypeScript());
});

gulp.task('clean-logs', function () {
    return del([
        'npm-debug*'
    ]);
});

gulp.task('clean', function() {
    sequence('clean-jsmap', 'clean-logs');
});
