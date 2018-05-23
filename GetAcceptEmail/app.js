lbs.apploader.register('GetAcceptEmail', function () {
    var self = this;

    /*Config (version 2.0)
        This is the setup of your app. Specify which data and resources that should loaded to set the enviroment of your app.
        App specific setup for your app to the config section here, i.e self.config.yourPropertiy:'foo'
        The variabels specified in "config:{}", when you initalize your app are available in in the object "appConfig".
    */
    self.config =  function(appConfig){
            this.yourPropertyDefinedWhenTheAppIsUsed = appConfig.yourProperty;
            this.dataSources = [];
            this.resources = {
                scripts: [], // <= External libs for your apps. Must be a file
                styles: ['app.css'], // <= Load styling for the app.
                libs: ['json2xml.js'] // <= Already included libs, put not loaded per default. Example json2xml.js
            };
    };

    //initialize
    /*Initialize
        Initialize happens after the data and recources are loaded but before the view is rendered.
        Here it is your job to implement the logic of your app, by attaching data and functions to 'viewModel' and then returning it
        The data you requested along with localization are delivered in the variable viewModel.
        You may make any modifications you please to it or replace is with a entirely new one before returning it.
        The returned viewModel will be used to build your app.
        
        Node is a reference to the HTML-node where the app is being initalized form. Frankly we do not know when you'll ever need it,
        but, well, here you have it.
    */
    self.initialize = function (node, viewModel) {
        var documentDataXml = lbs.common.executeVba("GetAccept.GetEmailData");

        try {
            var json = xml2json($.parseXML(documentDataXml), '');
            json = $.parseJSON(json);
        }
        catch(e) {
            alert(e);
        }

        if (!!json) {
            viewModel.emailSubject = ko.observable(unpackStringAfterSending(json.emailData.emailSubject));
            viewModel.emailMessage = ko.observable(unpackStringAfterSending(json.emailData.emailMessage));
        }
        
        function storeEmailData() {
            var message = prepereStringForSendingToVBA(viewModel.emailMessage());
            var subject = prepereStringForSendingToVBA(viewModel.emailSubject());

            var data = {
                'emailData': {
                    'emailMessage': message,
                    'emailSubject': subject
                }
            };
        
            try {
                var t = json2xml(data);
            }
            catch(e) {
                alert(e);
            }
            lbs.common.executeVba('GetAccept.StoreEmailData,' + t);

            window.open('', '_parent', '');
            window.close();
        }

        function unpackStringAfterSending(str) {
            str = str.replace(/%0/g,',');
            str = str.replace(/%1/g, String.fromCharCode(10));
            str = str.replace(/%2/g, String.fromCharCode(39));
            str = str.replace(/%3/g, String.fromCharCode(34));
            str = str.replace(/%4/g, '&');
            str = str.replace(/%5/g, '>');
            str = str.replace(/%6/g, '<');

            return str;
        }

        function cancel() {
            window.open('', '_parent', '');
            window.close();
        }


        function prepereStringForSendingToVBA(str) {
            try {
                str = str.replace(/,/g,'%0');
                str = str.replace(/(?:\r\n|\r|\n)/g,'%1');
                str = str.replace(/'/g,'%2');
                str = str.replace(/"/g,'%3');
                str = str.replace(/&/g,'%4');
                str = str.replace(/>/g,'%5');
                str = str.replace(/</g,'%6');
            }   
            catch(e) {
                alert(e);
            }
            return str; 
        }

        viewModel.cancel = cancel;
        viewModel.storeEmailData = storeEmailData;
        return viewModel;
    };
});
