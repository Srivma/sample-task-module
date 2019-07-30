(function(window){

    var WORKER_PATH = '/scripts/recorderWorker.js';
    var recorder = function(source, configure){
      var config = configure || {};
      var bufferLen = config.bufferLen || 4096;
      this.context = source.context;
      this.node = this.context.createScriptProcessor(bufferLen, 2, 2);
      var worker = new Worker(config.workerPath || WORKER_PATH);

      worker.postMessage({
        command: 'init',
        config: {
          sampleRate: this.context.sampleRate
        }
      });

      var recording = false,
      currCallback;
      // when we click on start recording button it will come to this function and start buffering on the recording data
      this.node.onaudioprocess = function(e){
        if (!recording) return;
        worker.postMessage({
          command: 'record',
          buffer: [
            e.inputBuffer.getChannelData(0),
            e.inputBuffer.getChannelData(1)
          ]
        });
      }

      this.configure = function(configure){
        for (var prop in configure){
          if (configure.hasOwnProperty(prop)){
            config[prop] = configure[prop];
          }
        }
      }

      this.record = function(){
      recording = true;
      }

      this.stop = function(){
        recording = false;
      }

      this.clear = function(){
        worker.postMessage({ command: 'clear' });
      }
      // get the audio file buffer data and send to postmessage
      this.getBuffer = function(cb) {
        currCallback = cb || config.callback;
        worker.postMessage({ command: 'getBuffer' })
      }
      // we called exportWAV function in stoprecording button click to get the blob data
      this.exportWAV = function(cb, type){
        currCallback = cb || config.callback;
        type = type || config.type || 'audio/wav';
        if (!currCallback) throw new Error('Callback not set');

        worker.postMessage({
          command: 'exportWAV',
          type: type
        });
      }

      worker.onmessage = function(e){
        var blob = e.data;
        currCallback(blob);
      }

      source.connect(this.node);
      this.node.connect(this.context.destination);    //this should not be necessary
    };
    // download the wav file 
    recorder.forceDownload = function(blob, filename){
      var url = (window.URL || window.webkitURL).createObjectURL(blob);
      var link = window.document.createElement('a');
      link.href = url;
      link.download = filename || 'output.wav';
      var click = document.createEvent("Event");
      click.initEvent("click", true, true);
      link.dispatchEvent(click);
    }

    window.recorder = recorder;

  })(window);

  