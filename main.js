const electron = require('electron');
const app = electron.app;
const BrowserWindow = electron.BrowserWindow;
let mainWindow;

const path = require('path');

process.on('uncaughtException', function (err) {
  app.quit();
});

// quit after close app
app.on('window-all-closed', function () {
  app.quit();
  console.log('server stopped')
});

app.on('ready', function () {
  // let python = require('child_process').spawn('python3', ['app.py']);
  let python = require('child_process').spawn('python', [path.join(__dirname, 'app.py')]);
  python.unref()

  const rq = require('request-promise');
  const mainAddr = 'http://localhost:5000/';

  const openWindow = function () {
    mainWindow = new BrowserWindow({
      width: 800,
      height: 600,
      webPreferences: {
        nodeIntegration: true,
      }
    });
    mainWindow.loadURL(mainAddr);

    // terminate process
    mainWindow.on('closed', function () {

      // remove cache
      electron.session.defaultSession.clearCache(() => {})
      mainWindow = null;
    });
  };

  const startUp = function () {
    rq(mainAddr)
      .then(function (htmlString) {
        console.log('server started');
        openWindow();
      })
      .catch(function (err) {
        // tmp_error = err;
        // if (err == tmp_error) {
        //   console.log('server error: ' + err);
        // }
        startUp();
      });
  };

  startUp();
});