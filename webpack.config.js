var webpack = require('webpack');
var minified = process.env.MINIFIED === 'true';

module.exports = {
  entry: './src/kurve.ts',
  output: {
    library: "Kurve",
    libraryTarget: "umd",
    filename: minified ? 'dist/kurve.min.js' : 'dist/kurve.js'
  },
  devtool: 'source-map',
  plugins: minified ? [ new webpack.optimize.UglifyJsPlugin({minimize: true}) ] : [],
  resolve: {
    extensions: ['', '.webpack.js', '.web.js', '.ts']
  },
  module: {
    loaders: [
      { loader: 'ts-loader' }
    ]
  }
}
