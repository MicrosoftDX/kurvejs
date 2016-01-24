var gulp = require("gulp");
var uglify = require("gulp-uglify");
var typescript = require("gulp-typescript");
var sourcemaps = require("gulp-sourcemaps");
var merge2 = require("merge2");
var concat = require("gulp-concat");
var rename = require("gulp-rename");
var changed = require('gulp-changed');
var runSequence = require('run-sequence');
var replace = require("gulp-replace");
var config = require("./config.json");
var addModuleExports = require("./gulp-addModuleExports");
/*
Compiles all typescript files and creating a declaration file.
*/
gulp.task('default', ['typescript-compile'], function () {
    gulp.src(config.build.srcOutputDirectory + '-' + config.build.version +  '/kurve-' + config.build.version + '.js')
    .pipe(addModuleExports("Kurve"))
    .pipe(gulp.dest(config.build.outputDirectory + '-' + config.build.version + '/'))
    .pipe(rename(config.build.minFilename + '-' + config.build.version + '.js'))
    .pipe(uglify())
    .pipe(gulp.dest(config.build.outputDirectory + '-' + config.build.version + '/'));
});


gulp.task('typescript-sourcemaps',['move-html'], function () {
    var tsResult = gulp.src(config.core.typescript) 
                     .pipe(sourcemaps.init()) // sourcemaps init. currently redundant directory def, waiting for this - https://github.com/floridoo/gulp-sourcemaps/issues/111 
                     .pipe(typescript({  
                             noExternalResolve: true,  
                             target: 'ES5',  
                             declarationFiles: true, 
                             typescript: require('typescript'),
                             out: 'kurve-' + config.build.version + '.js' 
                     })); 
     return tsResult.js 
             .pipe(sourcemaps.write("./")) // sourcemaps are written. 
             .pipe(gulp.dest(config.build.srcOutputDirectory + '-' + config.build.version + '/')); 

});

gulp.task('typescript-compile',["typescript-sourcemaps"], function() {   
       var tsResult = gulp.src(config.core.typescript) 
                     .pipe(typescript({  
                             noExternalResolve: true,  
                             target: 'ES5',  
                             declarationFiles: true, 
                             typescript: require('typescript'),
                             out: 'kurve-' + config.build.version + '.js' 
                     })); 
     return merge2([ 
         tsResult.dts 
             .pipe(concat('kurve-' + config.build.version  + config.build.declarationFilename)) 
             .pipe(gulp.dest(config.build.outputDirectory + '-' + config.build.version + '/')), 
         tsResult.js 
             .pipe(gulp.dest(config.build.srcOutputDirectory + '-' + config.build.version + '/')) 
     ]); 
 }); 

gulp.task('app-compile', ["typescript-compile"], function () {
    return gulp
        .src(['./Examples/OAuthV1/app.ts',
			  './Examples/OAuthV2/app.ts', 
			  './Examples/NoWindow/noLoginWindow.ts', 
			  './dist-' + config.build.version + '/kurve-' + config.build.version + '.d.ts'], 
			  { base: '.' })
        .pipe(typescript({
            noExternalResolve: true,
            target: 'ES5',
            outDir: './',
            typescript: require('typescript')
        }))
        .pipe(gulp.dest('./'));
});

/**
 * Watch task, will call the default task if a js file is updated.
 */
gulp.task('watch', function () {
    gulp.watch(config.core.typescript, ['default']);
});

gulp.task('move-html', function () {
    var result = gulp.src(['login.html'])
        .pipe(gulp.dest('./dist-' + config.build.version));
    return result;
});


