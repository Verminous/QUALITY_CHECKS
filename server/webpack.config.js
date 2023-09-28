module.exports = {
    // ... other config settings
    devServer: {
      setupMiddlewares(middlewares, devServerOptions) {
        // Your middleware logic here
        return middlewares;
      },
    },
  };