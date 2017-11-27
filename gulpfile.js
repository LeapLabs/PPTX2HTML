'use strict'

const browserify = require('browserify')
const uglifyES = require('uglify-es')
const gulp = require('gulp')
const buffer = require('vinyl-buffer')
const source = require('vinyl-source-stream')
const composer = require('gulp-uglify/composer')
const sourcemaps = require('gulp-sourcemaps')
const minify = composer(uglifyES, console)
const jetpack = require('fs-jetpack')
const rename = require('gulp-rename')
const babelify = require('babelify')
const babel = require('@babel/core')

const srcDir = jetpack.cwd('./src')

const distDir = jetpack.cwd('./dist')

gulp.task('clean-dist', () => {
  distDir.dir('.', {empty: true})
})

const buildJsFile = ({filePath, name}) => {
  const b = browserify({
    entries: filePath,
    debug: true, // sourcemaps
    standalone: name
  })

  b.transform(babelify.configure({
    presets: [
      ['@babel/preset-env', {
        useBuiltIns: 'usage'
      }]
    ],
    babel
  }))

  return b.bundle()
    .pipe(source(name + '.js'))
    .pipe(buffer())
    .pipe(sourcemaps.init({loadMaps: true}))
    .pipe(sourcemaps.write('.'))
    .pipe(gulp.dest(distDir.path()))
}

gulp.task('build-main', ['clean-dist'], () =>
  buildJsFile({filePath: srcDir.path('main.js'), name: 'pptx2html'})
)

gulp.task('build-worker', ['clean-dist', 'build-main'], () =>
  buildJsFile({filePath: srcDir.path('worker.js'), name: 'pptx2html_worker'})
)

gulp.task('minify', ['build-main', 'build-worker'], () => {
  return gulp.src(distDir.path('*.js'))
    .pipe(rename(p => { p.extname = '.min.js' }))
    .pipe(buffer())
    .pipe(sourcemaps.init({loadMaps: true}))
    .pipe(minify({
      compress: {
        passes: 2,
        typeofs: false
      },
      ie8: true
    }))
    .pipe(sourcemaps.write('.'))
    .pipe(gulp.dest(distDir.path()))
})
