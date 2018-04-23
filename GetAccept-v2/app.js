lbs.apploader.register('GetAccept-v2', function () {
    var self = this;

    /*Config (version 2.0)
        This is the setup of your app. Specify which data and resources that should loaded to set the enviroment of your app.
        App specific setup for your app to the config section here, i.e self.config.yourPropertiy:'foo'
        The variabels specified in "config:{}", when you initalize your app are available in in the object "appConfig".
    */
    self.config = function (appConfig) {
        this.yourPropertyDefinedWhenTheAppIsUsed = appConfig.yourProperty;
        this.dataSources = [{
            type: 'activeInspector'
        }];
        this.appConfig = appConfig;
        this.resources = {
            scripts: ['pusher.min.js'], // <= External libs for your apps. Must be a file
            styles: ['app.css', 'animate.css'], // <= Load styling for the app.
            libs: ['json2xml.js'] // <= Already included libs, put not loaded per default. Example json2xml.js
        };
    };

    self.initialize = function (node, viewModel) {
        var appConfig = self.config.appConfig;
        var authEndpoint = "https://app.getaccept.com/api";
        var apiEndpoint = "https://app.getaccept.com/api";
        var clientId = "Lime CRM";
        var className = lbs.limeDataConnection.ActiveInspector.class.name;
        var class_id = 'id' + className;
        var originalPersonList = [];
        var originalTemplateList = [];
        var refreshToken;
        var entityId;
        var userHash;
        var tokenHandler = "";
        var accessToken = "";
        var videoTimer = false;
        var pusherInit = false;

        var id = lbs.limeDataConnection.ActiveInspector.Record.id;
        var key = className + id;

        viewModel.searchToInvite = true;
        viewModel.User = "";

        // THIS IS THE DIFERENT VIEWS
        viewModel.Spinner = ko.observable(false);
        viewModel.Signup = ko.observable(false);
        viewModel.Login = ko.observable(true);
        viewModel.Settings = ko.observable(false);
        viewModel.GaDocuments = ko.observable(false);
        viewModel.Recipient = ko.observable(false);
        viewModel.Document = ko.observable(false);
        viewModel.Video = ko.observable(false);
        viewModel.Reminder = ko.observable(false);
        viewModel.Invite = ko.observable(false);
        viewModel.Help = ko.observable(false);
        viewModel.Menu = ko.observable(false);
        viewModel.Analytics = ko.observable(false);
        viewModel.File = ko.observable(false);
        viewModel.TemplateFields = ko.observable(false);
        viewModel.AvailableFields = ko.observable(false);

        //Variabler
        viewModel.userName = ko.observable('');
        viewModel.password = ko.observable('');
        viewModel.tryingToLogIn = ko.observable(false);
        viewModel.showPersons = ko.observable(false);
        viewModel.searchValue = ko.observable('').extend({
            throttle: 500
        });
        viewModel.inviteSearchValue = ko.observable('').extend({
            throttle: 500
        });
        viewModel.templateSerach = ko.observable('').extend({
            throttle: 500
        });
        viewModel.smartReminders = ko.observable(false);
        viewModel.sendSMS = ko.observable(false);
        viewModel.useTemplates = ko.observable(false);
        viewModel.hasVideo = ko.observable(false);
        viewModel.sendInternal = ko.observable(false);
        viewModel.videoData = ko.observable({});
        viewModel.showApp = false;
        viewModel.emailSubject = ko.observable('');
        viewModel.emailMessage = ko.observable('');

        viewModel.documentAnalytics = ko.observable();
        viewModel.documentName = ko.observable('');
        //FREE ACCOUNT 
        viewModel.signupCountry = ko.observable('');
        viewModel.signupCompany = ko.observable('');
        viewModel.signupEmail = ko.observable('');
        viewModel.signupName = ko.observable('');
        viewModel.signupPassword = ko.observable('')
        viewModel.signupCountryList = ko.observableArray();

        //TEMPLATE 
        viewModel.selectedTemplate = ko.observable();
        viewModel.selectedTemplateId = ko.observable();
        viewModel.selectedTemplateFields = ko.observableArray();
        viewModel.availableLimeFields = ko.observableArray();

        //LISTS
        viewModel.document = ko.observableArray();
        viewModel.gaDocumentList = ko.observableArray();
        viewModel.documentList = ko.observableArray();

        viewModel.uploadedDocuments = ko.observableArray();
        viewModel.personList = ko.observableArray();
        viewModel.recipientsList = ko.observableArray();
        viewModel.worksteps = ko.observableArray();
        viewModel.entityList = ko.observableArray();
        viewModel.coworkerList = ko.observableArray();
        viewModel.templateList = ko.observableArray();
        viewModel.reminderOptions = [1, 3, 5, 7, 10, 14];

        function toogleGaView() {
            $('.ga-container').slideToggle("slow");
            if ($('.ga-container').hasClass("extended")) {
                $('.ga-container').removeClass("extended");
                $('.chevron').removeClass("fa-chevron-up").addClass("fa-chevron-down");
                lbs.bakery.setCookie("shouldToggle", null, -1);
            } else {
                $('.ga-container').addClass("extended");
                $('.chevron').removeClass("fa-chevron-down").addClass("fa-chevron-up");
                lbs.bakery.setCookie("shouldToggle", true, 30);
            }
        }

        function logon() {
            if (viewModel.userName().length > 0 && viewModel.password().length > 0) {
                var email = viewModel.userName();
                var password = viewModel.password();
                var postUrl = authEndpoint + "/v1/auth";

                viewModel.tryingToLogIn(true);
                var xhr = new XMLHttpRequest();
                xhr.open('POST', postUrl, true);
                xhr.setRequestHeader('Content-type', 'application/json');
                xhr.onreadystatechange = function () {
                    if (xhr.readyState == 4) {
                        status = xhr.status;
                        if (status == 200) {
                            viewModel.tryingToLogIn(false);
                            data = JSON.parse(xhr.responseText);
                            saveToken(data);
                            loadUserSettings();
                        } else {
                            data = JSON.parse(xhr.responseText);
                            if (!data.error) {
                                viewModel.tryingToLogIn(false);
                                alert(viewModel.localize.GetAccept.VERIFY_CREDENTIALS);
                            } else {
                                viewModel.tryingToLogIn(false);
                                alert(data.error);
                            }
                            return false;
                        }
                    }
                }
                var json = '{ "email": "' + email + '","password": "' + password + '", "client_id": "' + clientId + '" }';
                xhr.send(json);
            }
        }

        function loadUserSettings() {
            apiRequest("users/me", "GET", "", function (data) {
                if (!!data.user) {
                    var userHash = data.user.id;
                    lbs.bakery.setCookie("userHash", userHash, 30);
                    viewModel.Login(false);
                    viewModel.GaDocuments(true);
                    initGa();
                }
            });
        }

        function backToLogin() {
            viewModel.Login(true);
            viewModel.Signup(false);
        }

        function logout() {
            accessToken = "";
            refreshToken = "";
            expireToken = "";
            lbs.bakery.setCookie("accessToken", null, -1);
            lbs.bakery.setCookie("refreshToken", null, -1);
            lbs.bakery.setCookie("expireToken", null, -1);
            lbs.bakery.setCookie("entityId", null, -1);
            lbs.bakery.setCookie("userHash", null, -1);

            $('.win-document').addClass('hidden');
            $('.win-auth').removeClass('hidden');
            lbs.common.executeVba("GetAccept.SetTokens", "-");
            cancel();
            viewModel.GaDocuments(false);
            viewModel.Login(true);
            initGa();
        }

        function saveToken(data) {
            if (data.access_token) {
                accessToken = data.access_token;
            }
            if (data.refresh_token) {
                refreshToken = data.refresh_token;
            }
            if (data.entity_id) {
                entityId = data.entity_id;
                lbs.bakery.setCookie("entityId", entityId, 30);
            }
            if (data.user_hash) {
                userHash = data.user_hash;
                lbs.bakery.setCookie("userHash", userHash, 30);
            }
            var expireToken = Math.ceil(new Date().getTime() / 1000) + data.expires_in;
            lbs.bakery.setCookie("accessToken", accessToken, 30);
            lbs.bakery.setCookie("refreshToken", refreshToken, 30);
            lbs.bakery.setCookie("expireToken", expireToken, 30);
            lbs.bakery.setCookie("fullToken", JSON.stringify(data), 30);
            return;
        }


        function checkLogin() {
            var have_token = false;
            if (typeof lbs.bakery.getCookie("accessToken") != 'undefined') {
                if (lbs.bakery.getCookie("accessToken") != '') {
                    have_token = true;

                }
            }
            if (have_token) {
                accessToken = lbs.bakery.getCookie("accessToken");
                apiRequest("test", "GET", "", function (data) {
                    if (data) {
                        refreshToken = lbs.bakery.getCookie("refreshToken");
                        expireToken = parseInt(lbs.bakery.getCookie("expireToken"));
                        entityId = lbs.bakery.getCookie("entityId");
                        userHash = lbs.bakery.getCookie("userHash");
                        fullToken = lbs.bakery.getCookie("fullToken");
                        if (fullToken) {
                            lbs.common.executeVba("GetAccept.SetTokens", fullToken);
                        }
                        var nowSec = Math.ceil(new Date().getTime() / 1000);

                        //Check if token expires within 7 days
                        if (expireToken - nowSec > 604800) {
                            var validTo = moment(expireToken * 1000).format("YYYY-MM-DD hh:mm a");
                            console.log("Token expires: " + validTo.toString());
                        } else {
                            apiRequest("refresh", "GET", "", function (data) {
                                saveToken(data);
                            });
                        }

                        viewModel.Login(false);
                        listCoworkers();
                        listEntities();
                        listDocuments();
                        setupPusher();

                        if (!!lbs.bakery.getCookie("shouldToggle") && lbs.bakery.getCookie("shouldToggle") !== 'undefined') {
                            setTimeout(function () {
                                toogleGaView();
                            }, 500);
                        }

                    } else {
                        viewModel.Login(true);
                    }
                });


            } else {
                return false;
            }
        }

        function documentAnalytics(documentData) {
            if (!!documentData.id) {
                viewModel.Spinner(true);
                apiRequest("documents/" + documentData.id + "?with_pages=true&with_stats=true", "GET", "", function (data) {
                    if (!!data.send_date) {
                        var analyticsData = {};
                        analyticsData.pageCount = data.stats.document_page_completion + "/" + data.stats.document_page_count;
                        analyticsData.name = data.name;
                        analyticsData.status = data.status.toLowerCase();
                        var date = moment(data.send_date).fromNow(true);
                        analyticsData.sendDate = date.substr(0, date.indexOf(' '));
                        analyticsData.sendDateText = date.substr(date.indexOf(' ') + 1);
                        analyticsData.visits = data.stats.document_visit_count;
                        var visitTime = sec2time(data.stats.document_visit_time);
                        analyticsData.visitTime = visitTime;
                        analyticsData.url = documentData.sso_url;

                        analyticsData.delete = function () {
                            apiRequest("documents/" + documentData.id, "DELETE", "", function (data) {
                                cancel();
                                refreshGaDocuments();
                            });
                        }

                        analyticsData.analyticPageList = [];

                        _.each(data.pages, function (page) {
                            page.timespent = sec2str(page.page_time)
                            var percent = Math.round((page.page_time / data.stats.document_visit_time) * 100);
                            page.percent = !!percent ? percent : 0;
                            analyticsData.analyticPageList.push(page);
                        });

                        viewModel.documentAnalytics(analyticsData);

                        viewModel.Analytics(true);
                        viewModel.GaDocuments(false);
                        viewModel.Spinner(false);
                    } else {
                        if (documentData.sso_url) {
                            lbs.common.executeVba('shell,' + documentData.sso_url);
                        }
                        viewModel.Spinner(false);
                    }
                });
            }
        }

        function sec2time(sec) {
            var seconds = sec * 1000;
            var min = Math.round(moment.duration(seconds).asMinutes());
            if (min > 0) {
                var hours = Math.round(moment.duration(seconds).asHours());
                if (hours > 0) {
                    var days = Math.round(moment.duration(seconds).asDays());
                    if (days > 0) {
                        return {
                            "type": "days",
                            "value": days
                        }
                    } else {
                        return {
                            "type": "hours",
                            "value": hours
                        }

                    }
                } else {
                    return {
                        "type": "min",
                        "value": min
                    }
                }
            } else {
                return {
                    "type": "sec",
                    "value": sec
                }
            }
        }

        function sec2str(t) {
            if (typeof t != 'undefined') {
                var d = Math.floor(t / 86400),
                    h = ('' + Math.floor(t / 3600) % 24).slice(-2),
                    m = ('' + Math.floor(t / 60) % 60).slice(-2),
                    s = Math.round(t % 60);
                var str = (d > 0 ? d + 'd ' : '') + (h > 0 ? h + 'h ' : '') + (m > 0 ? m + 'm ' : '') + (t < 1 ? '' : s + 's');
                return str ? str : '0s';
            } else {
                return '0s';
            }
        }

        function listDocuments() {
            viewModel.Spinner(true);
            viewModel.gaDocumentList.removeAll();
            viewModel.GaDocuments(true);
            var active_record_id = lbs.limeDataConnection.ActiveInspector.Controls.GetValue(class_id);
            if (!!active_record_id) {
                var document_ids = lbs.common.executeVba("GetAccept.CheckDocuments," + active_record_id + ',' + className);
                if (document_ids && document_ids !== "False") {
                    apiRequest('documents?external_id=' + document_ids, 'GET', '', function (data) {
                        if (data) {
                            processDocument(data);
                        } else {
                            console.log("Couldn't get document data");
                        }
                    });
                } else {
                    viewModel.Spinner(false);
                    console.log("Couldn't find any document ids.");
                }

                apiRequest('documents?external_id=' + key, 'GET', '', function (data) {
                    if (data) {
                        processDocument(data);
                    }
                });

                apiRequest('entity', 'GET', '', function (data) {
                    viewModel.emailSubject(data.entity.email_send_subject === '' ? data.entity.default_email_send_subject : data.entity.email_send_subject);
                    viewModel.emailMessage(data.entity.email_send_message === '' ? data.entity.default_email_send_message : data.entity.email_send_message);
                });

            } else {
                className = lbs.limeDataConnection.ActiveInspector.class.name;
                class_id = 'id' + className;
                viewModel.Spinner(false);
                listDocuments();
                console.log("Couldn't find active_record_id. Restarts the loading of documents");
            }


        }

        function processDocument(data) {
            viewModel.Spinner(false);
            if (data.length > 0) {
                $.each(data, function (index, doc) {
                    var docUrl = '/document/' + (doc.status == 'draft' ? 'edit' : 'view') + '/' + doc.id;
                    var sso_url = 'https://app.getaccept.com/auth/sso/login?token=' + escape(accessToken) + '&entity_id=' + entityId + '&go=' + escape(docUrl);
                    var document = {
                        id: doc.id,
                        name: doc.name,
                        status: doc.status,
                        sso_url: sso_url,
                        is_signing: doc.is_signing
                    }
                    document.analytics = function () {
                        documentAnalytics(this);
                    }

                    document.download = function () {
                        downloadDocument(this);
                    }

                    viewModel.gaDocumentList.push(document);
                });
            }
        }

        function downloadDocument(document) {
            var document_id = document.id;
            var documentname = prepereStringForSavingToFileSystem(document.name);
            //Use parameter direct,  /download?direct=true to get binary content back
            //This can later be processed in VBA to store file in Lime
            apiRequest('documents/' + document_id + '/download', 'GET', '', function (data) {
                if (typeof (data.document_url) != 'undefined') {
                    alert(viewModel.localize.GetAccept.DOCUMENT_IS_DOWNLOADED);
                    lbs.common.executeVba("GetAccept.DownloadFile," + data.document_url + ',' + documentname + ',' + className + ',' + appConfig.title_field);
                } else {
                    alert('Could not find signed document');
                }
            });
        }

        function apiRequest(action, method, json, callback) {
            var postUrl = apiEndpoint + "/v1/" + action;
            var xhr = new XMLHttpRequest();
            xhr.open(method, postUrl, true);
            xhr.setRequestHeader('Content-type', 'application/json');
            if (accessToken) {
                xhr.setRequestHeader('Authorization', 'bearer ' + accessToken);
            }
            xhr.onreadystatechange = function () {
                // If the request completed
                if (xhr.readyState == 4) {
                    status = xhr.status;
                    if (status == 200) {
                        callback(JSON.parse(xhr.responseText));
                    } else if (status == 401) {
                        callback(false);
                        //alert('Something wrong! Try to re-login.');
                    } else {
                        callback(false);
                    }
                }
            };
            if (json) {
                xhr.send(JSON.stringify(json));
            } else {
                xhr.send();
            }
        }

        function showRecipientPicker() {
            viewModel.Spinner(true);
            viewModel.worksteps([1, 0]);
            hideAllSteps();
            getRecipients();
            viewModel.Recipient(true);
            viewModel.GaDocuments(false);
            viewModel.Spinner(false);
            if (!$('.ga-container').hasClass("extended")) {
                toogleGaView();
            }
        }

        function cancel() {
            viewModel.worksteps([]);
            hideAllSteps();
            viewModel.GaDocuments(true);
            viewModel.documentList.removeAll();
            viewModel.recipientsList.removeAll();
            viewModel.personList.removeAll();
            viewModel.document.removeAll();
        }

        function done() {
            hideAllSteps();
            viewModel.documentList.removeAll();
            viewModel.recipientsList.removeAll();
            viewModel.personList.removeAll();
            setTimeout(function () {
                viewModel.GaDocuments(true);
                listDocuments();
            }, 1000)
        }

        function backDocument() {
            viewModel.worksteps([1, 0]);
            hideAllSteps();
            viewModel.File(true);
            viewModel.documentList.removeAll();
            viewModel.document.removeAll();
        }

        function backFile() {
            viewModel.worksteps([1, 0]);
            hideAllSteps();
            viewModel.Recipient(true);
            viewModel.documentList.removeAll();
            viewModel.document.removeAll();
        }

        function backFromVideo() {
            viewModel.Document(true);
            viewModel.Video(false);
        }

        function showSettings() {
            hideAllSteps();
            viewModel.Settings(true);
        }

        function showInvite() {
            hideAllSteps();
            viewModel.Invite(true);
        }

        function showHelp() {
            hideAllSteps();
            viewModel.Help(true);
        }

        function showMenu() {
            viewModel.Menu(!viewModel.Menu());
        }

        function hideMenu() {
            viewModel.Menu(false);
        }

        function uploadVideo() {
            viewModel.Video(true);
        }

        function removeVideo() {
            viewModel.videoData({});
            viewModel.hasVideo(false);
        }

        function closeVideo() {
            viewModel.Video(false);
        }

        function createFreeAccount() {
            var countries = listCounties();
            viewModel.signupCountryList.removeAll();
            _.each(countries, function (name, key) {
                var country = {};
                country.key = key;
                country.name = name;

                country.selectCountry = function () {
                    viewModel.signupCountry(this);
                }
                viewModel.signupCountryList.push(country);
            });

            viewModel.signupCountry(viewModel.signupCountryList()[0]);

            var email = '';
            var name = '';
            var company = '';
            try {
                name = lbs.limeDataConnection.ActiveUser.Record("name");
                email = lbs.limeDataConnection.ActiveUser.Record("email");
                company = lbs.limeDataConnection.Database.Name;
            } catch (e) {
                console.log(e);
            }

            viewModel.signupCompany(company);
            viewModel.signupEmail(email);
            viewModel.signupName(name);
            viewModel.signupPassword('');

            viewModel.Login(false);
            viewModel.Signup(true);
        }

        function registerSignup() {
            if (confirm("You are about to create a new GetAccept account.")) {
                var mobile = '';
                var first_name = '';
                var last_name = '';
                try {
                    mobile = lbs.limeDataConnection.ActiveUser.Record("cellphone");
                    first_name = lbs.limeDataConnection.ActiveUser.Record("firstname");
                    last_name = lbs.limeDataConnection.ActiveUser.Record("lastname");
                } catch (e) {
                    console.log(e);
                }

                if (viewModel.signupName() !== '' && viewModel.signupCompany() !== '' && viewModel.signupEmail() !== '' && viewModel.signupPassword() !== '') {
                    var json = {
                        user_registration_source: "Lime CRM",
                        client_id: "Lime CRM",
                        auto_login: true,
                        skip_invitation: true,
                        email: viewModel.signupEmail(),
                        password: viewModel.signupPassword(),
                        first_name: first_name,
                        last_name: last_name,
                        full_name: viewModel.signupName(),
                        mobile: mobile,
                        entity_country_code: viewModel.signupCountry().key,
                        entity_name: viewModel.signupCompany()
                    }

                    apiRequest('register', 'POST', json, function (data) {
                        if (data.error) {
                            alert(data.error);
                        } else {
                            viewModel.Signup(false);
                            viewModel.Login(true);
                            viewModel.userName(viewModel.signupEmail());
                        }
                    });

                } else {
                    alert("You need to fill in all fields.")
                }
            }
        }

        function hideAllSteps() {
            viewModel.Document(false);
            viewModel.Recipient(false);
            viewModel.GaDocuments(false);
            viewModel.Settings(false);
            viewModel.Invite(false);
            viewModel.Help(false);
            viewModel.Reminder(false);
            viewModel.Menu(false);
            viewModel.Signup(false);
            viewModel.Analytics(false);
            viewModel.File(false);
            viewModel.TemplateFields(false);
            viewModel.AvailableFields(false);
        }

        function getRecipients() {
            viewModel.personList.removeAll();
            var contactString = lbs.common.executeVba("GetAccept.GetContactList," + className);
            var contacts = !!contactString ? JSON.parse(contactString) : false;
            originalPersonList = [];
            if (contacts) {
                $.each(contacts.Persons, function (index, personData) {
                    if (validateEmail(personData.email)) {
                        var person = new recipientModel(personData, false);
                        originalPersonList.push(person);
                        viewModel.personList.push(person);
                    }
                });

                /*var coworkers = lbs.common.executeVba('GetAccept.GetCoworkerList');
                coworkers = JSON.parse(coworkers);
                _.each(coworkers.Persons, function (coworkerData) {
                    if (coworkerData.email !== "") {
                        var coworker = new recipientModel(coworkerData, true);
                        originalPersonList.push(coworker);
                        viewModel.personList.push(coworker);
                    }
                });*/
            }
        }

        function getTemplates() {
            viewModel.templateList.removeAll();
            originalTemplateList = [];
            apiRequest('templates', 'GET', '', function (data) {
                if (data.templates) {
                    $.each(data.templates, function (index, templateData) {
                        var temp = new templateModel(templateData);
                        viewModel.templateList.push(temp);
                        originalTemplateList.push(temp);
                    });
                }
            });
        }

        function getTemplateFields() {
            viewModel.selectedTemplateFields.removeAll();
            apiRequest("templates/" + viewModel.selectedTemplateId() + '/fields', "GET", "", function (data) {
                if (data) {
                    $.each(data.fields, function (index, fieldData) {
                        if (!!fieldData.field_label) {
                            try {
                                var fieldString = fieldData.field_value;
                                var fieldKey = fieldString.replace("{{", "").replace("}}", "");
                                var fieldKeyValue = eval('viewModel.' + fieldKey + '.text');
                                fieldData.field_value = !!fieldKeyValue ? fieldKeyValue : field_value;
                            } catch (e) {
                                console.log(e);
                            }
                            viewModel.selectedTemplateFields.push(fieldData);
                        }
                    });
                    viewModel.TemplateFields(true);
                } else {
                    viewModel.Document(true);
                }
            });
        }

        function showTemplateParameters() {
            hideAllSteps();
            viewModel.availableLimeFields.removeAll();
            $.each(eval('viewModel.' + className), function (fieldName, fieldValue) {
                var fieldData = new mapAvailableField(fieldName, fieldValue);
                viewModel.availableLimeFields.push(fieldData);
            });
            viewModel.AvailableFields(true);
        }

        function mapAvailableField(fieldName, fieldValue) {
            var field = this;
            try {
                field.value = fieldValue.text;
                field.name = lbs.limeDataConnection.ActiveInspector.Record.Field(fieldName).LocalName
                field.key = "{{" + className + "." + fieldName + "}}";
                field.copy = function () {
                    try {
                        window.clipboardData.setData('Text', this.key);
                        alert("Value is copied to clipboard");
                    } catch (err) {
                        console.log(err);
                    }
                }
            } catch (e) {
                console.log(e);
            }
            return field;
        }

        function templateModel(templateData) {
            var template = this;
            template.name = templateData.name;
            template.id = templateData.id;
            template.thumb = templateData.thumb_url;
            template.selectTemplate = function () {
                viewModel.selectedTemplate(this);
                viewModel.selectedTemplateId(this.id);
            }
            return template;
        }

        function enterPress(d, e) {
            if (e.which == 66 && e.ctrlKey) {
                //catches ctrl + b
                return false;
            }
        }
        viewModel.enterPress = enterPress;

        function showPersonList() {
            if (!viewModel.sendInternal()) {
                getRecipients();
            }
            viewModel.showPersons(true);
        }

        function hidePersonList() {
            viewModel.showPersons(false);
        }

        function createTodo() {
            lbs.common.executeVba('GetAccept.CreateTodo, ' + this);
            done();
        }

        function recipientModel(personData, internal) {
            recipient = this;
            if (personData) {
                try {
                    recipient.name = personData.firstname + ' ' + personData.lastname;
                    recipient.firstname = personData.firstname;
                    recipient.lastname = personData.lastname;
                    recipient.email = personData.email;
                    recipient.internal = internal;
                    recipient.mobilephone = personData.mobilephone;
                    recipient.signer = ko.observable(true);
                    recipient.cc = ko.observable(false);
                    recipient.searchString = (personData.firstname + '' + personData.lastname + '' + personData.email).toLocaleLowerCase();
                } catch (e) {
                    alert(e);
                }

                recipient.remove = function () {
                    var index = viewModel.recipientsList().indexOf(this);
                    if (index > -1) {
                        viewModel.personList.push(this);
                        viewModel.recipientsList.remove(this);
                    }
                }
                recipient.add = function () {
                    if (viewModel.recipientsList().indexOf(this) == -1) {
                        viewModel.personList.remove(this);
                        viewModel.recipientsList.push(this);
                    }
                    hidePersonList();
                }
                recipient.isCC = function () {
                    this.signer(false);
                    this.cc(!this.cc());
                }
                recipient.isSigner = function () {
                    this.signer(!this.signer());
                    this.cc(false);
                }
            }

            return recipient;
        }

        function gaRecipient(recipient) {
            return {
                email: recipient.email,
                first_name: recipient.firstname,
                last_name: recipient.lastname,
                mobile: recipient.mobilephone,
                role: recipient.signer() ? 'signer' : 'cc'
            }
        }

        function validateEmail(email) {
            var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
            var reg = new RegExp("\å|\ä|\ö");
            if (!reg.test(email)) {
                return re.test(email);
            } else {
                return false;
            }

        }

        viewModel.inviteSearchValue.subscribe(function (inputString) {
            if (inputString.length > 3) {
                viewModel.coworkerList.removeAll();
                viewModel.Spinner(true);
                try {
                    var coworkers = JSON.parse(lbs.common.executeVba("GetAccept.SearchCoworkerByEmail," + inputString));
                    if (coworkers.Persons) {
                        $.each(coworkers.Persons, function (i, coworker) {
                            var person = new coworkerModel(coworker, true);
                            viewModel.coworkerList.push(person);
                        });
                    }

                } catch (error) {
                    alert(error);
                }
                viewModel.Spinner(false);
            }
        });

        viewModel.searchValue.subscribe(function (inputString) {
            if (inputString.length > 3) {
                viewModel.personList.removeAll();
                inputString = (inputString.toLowerCase()).replace(' ', '');

                if (viewModel.sendInternal()) {
                    viewModel.Spinner(true);
                    try {
                        var coworkers = JSON.parse(lbs.common.executeVba("GetAccept.SearchCoworkerByEmail," + inputString));
                        var reg = new RegExp("\å|\ä|\ö");
                        if (coworkers.Persons) {
                            $.each(coworkers.Persons, function (i, coworker) {
                                if (!reg.test(coworker.email)) {
                                    var person = new recipientModel(coworker, true);
                                    viewModel.personList.push(person);
                                }
                            });
                        }
                    } catch (error) {
                        alert(error);
                    }
                    viewModel.Spinner(false);
                } else {
                    viewModel.Spinner(true);
                    $.each(originalPersonList, function (index, person) {
                        if (person.searchString.indexOf(inputString) > -1) {
                            if (viewModel.personList().indexOf(person) == -1) {
                                viewModel.personList.push(person);
                            }
                        }
                    });
                    viewModel.Spinner(false);
                }
            }
        });

        viewModel.templateSerach.subscribe(function (inputString) {
            if (inputString.length > 2) {
                viewModel.templateList.removeAll();
                viewModel.Spinner(true);
                $.each(originalTemplateList, function (index, template) {
                    if (template.name.indexOf(inputString) > -1) {
                        if (viewModel.templateList().indexOf(template) == -1) {
                            viewModel.templateList.push(template);
                        }
                    }
                });
                viewModel.Spinner(false);

            } else if (inputString.length == 0) {
                viewModel.templateList.removeAll();
                $.each(originalTemplateList, function (index, temp) {
                    viewModel.templateList.push(temp);
                });
            }
        });

        function showFile() {
            viewModel.Spinner(true);
            if (viewModel.recipientsList().length > 0) {
                getTemplates();
                viewModel.Recipient(false);
                viewModel.File(true);
                viewModel.GaDocuments(false);
                viewModel.Spinner(false);
                reloadFileName();

            } else {
                viewModel.Spinner(false);
                alert("You need to select at least one recipient!");
            }
        }

        function reloadFileName() {
            viewModel.documentName('');
            if (lbs.limeDataConnection.ActiveInspector.ActiveExplorer.class.name === "document") {
                if (lbs.limeDataConnection.ActiveInspector.ActiveExplorer.Selection.Count > 0) {
                    var document_data = lbs.common.executeVba("GetAccept.GetDocumentData," + className);
                    document_data = JSON.parse(document_data);
                    viewModel.documentName(document_data[0].file_name);
                }
            }
        }

        function toggleTemplatePicker() {
            viewModel.useTemplates(true);
            viewModel.useFile(false);
        }

        function toggleDocumentPicker() {
            viewModel.useTemplates(false);
            viewModel.useFile(true);
        }

        viewModel.useFile = ko.observable(true);

        viewModel.toggleDocumentPicker = toggleDocumentPicker;
        viewModel.toggleTemplatePicker = toggleTemplatePicker;
        viewModel.reloadFileName = reloadFileName;

        function goToDocumentAfterFields() {
            hideAllSteps();
            viewModel.Document(true);
        }
        viewModel.goToDocumentAfterFields = goToDocumentAfterFields;

        function showDocument() {
            if (!!viewModel.selectedTemplate() && viewModel.useTemplates()) {
                viewModel.worksteps([1, 1]);
                viewModel.File(false);
                viewModel.Recipient(false);
                viewModel.TemplateFields(false);
                viewModel.GaDocuments(false);
                getTemplateFields();
            } else if (viewModel.useFile()) {
                if (lbs.limeDataConnection.ActiveInspector.ActiveExplorer.class.name === "document") {
                    if (lbs.limeDataConnection.ActiveInspector.ActiveExplorer.Selection.Count > 0) {
                        viewModel.worksteps([1, 1]);
                        viewModel.File(false);
                        viewModel.Recipient(false);
                        viewModel.Document(true);
                        viewModel.GaDocuments(false);
                        getDocuments();
                        uploadDocument();
                    } else {
                        viewModel.Spinner(false);
                        alert("You must select at least one document.")
                    }
                } else {
                    viewModel.Spinner(false);
                    alert("Please select a document or a template step forward.")
                }
            } else {
                alert("Please select a document or a template step forward.")
            }


        }

        function getDocuments() {
            //Gets the selected documents and stores them in two variables. One that will be used as the GA document and one that will hold all file data as attachments.
            viewModel.document.removeAll();
            viewModel.documentList.removeAll();
            var documentData = JSON.parse(lbs.common.executeVba("GetAccept.GetDocuments," + className));
            if (documentData.length > 0) {
                $.each(documentData, function (i, document) {
                    var document_json = {
                        name: document.name,
                        external_id: document.id
                    };
                    var doc = new documentModel(document_json);
                    viewModel.documentList.push(doc);
                });
                var doc_josn = {
                    name: documentData[0].name,
                    external_id: documentData[0].id
                };
                var doc = new documentModel(doc_josn);
                viewModel.document.push(doc);
            }
        }

        function documentModel(data) {
            var document = this;
            var document_name = data.name.replace(/\.[^/.]+$/, "");
            document_name = document_name.replace(/\.|\_|\-|\?/g, ' ');
            if (document_name === document_name.toLowerCase()) {
                document_name = document_name.replace(/\b[a-zA-Z\u00C0-\u00ff]/g, function (letter) {
                    return letter.toUpperCase();
                });
            }
            document.documentName = ko.observable(document_name);
            document.external_id = data.external_id;
            return document;
        }

        function send() {
            sendDocument(true);
        }

        function open() {
            sendDocument(false);
        }

        function refreshGaDocuments() {
            listDocuments();
        }

        function packEmailData() {
            var emailData = {};
            try {
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
                } catch (e) {
                    alert(e);
                }

                lbs.common.executeVba("GetAccept.showEmailDialog," + t);
                var emailResult = lbs.common.executeVba("GetAccept.GetEmailData");

                unpackEmailData(emailResult);
            } catch (e) {
                alert(e);
            }
        }

        viewModel.updateMessage = ko.observable(false);
        viewModel.updateSubject = ko.observable(false);

        function unpackEmailData(emailResult) {
            try {
                var json = xml2json($.parseXML(emailResult), '');
                json = $.parseJSON(json);

                var message = unpackStringAfterSending(json.emailData.emailMessage);
                var subject = unpackStringAfterSending(json.emailData.emailSubject);

                if (message != viewModel.emailMessage()) {
                    viewModel.updateMessage(true);
                    setTimeout(function () {
                        viewModel.updateMessage(false);
                    }, 1000);
                    viewModel.emailMessage(message);
                }

                if (subject != viewModel.emailSubject()) {
                    viewModel.updateSubject(true);
                    setTimeout(function () {
                        viewModel.updateSubject(false);
                    }, 1000);
                    viewModel.emailSubject(subject);
                }
            } catch (e) {
                alert(e);
            }
        }

        function unpackStringAfterSending(str) {
            if (!!str) {
                str = str.replace(/%0/g, ',');
                str = str.replace(/%1/g, String.fromCharCode(10));
                str = str.replace(/%2/g, String.fromCharCode(39));
                str = str.replace(/%3/g, String.fromCharCode(34));
                str = str.replace(/%4/g, '&');
                str = str.replace(/%5/g, '>');
                str = str.replace(/%6/g, '<');
            }

            return str;
        }

        function prepereStringForSendingToVBA(str) {
            try {
                str = str.replace(/,/g, '%0');
                str = str.replace(/(?:\r\n|\r|\n)/g, '%1');
                str = str.replace(/'/g, '%2');
                str = str.replace(/"/g, '%3');
                str = str.replace(/&/g, '%4');
                str = str.replace(/>/g, '%5');
                str = str.replace(/</g, '%6');


            } catch (e) {
                alert(e);
            }
            return str;
        }

        function prepereStringForSavingToFileSystem(str) {
            try {
                //A filename cannot contain any of the following characters: \ / : * ? " < > |  
                str = str.replace(/,/g, '.');
                str = str.replace(/\//g, '');
                str = str.replace(/\\/g, '');
                str = str.replace(/\?/g, '');
                str = str.replace(/'/g, '');
                str = str.replace(/:/g, '');
                str = str.replace(/\|/g, '');
                str = str.replace(/\*/g, '');
                str = str.replace(/"/g, '');
                str = str.replace(/>/g, '');
                str = str.replace(/</g, '');
            } catch (e) {
                alert(e);
            }
            return str;
        }

        viewModel.createEmail = packEmailData;

        function sendDocument(automaticSending) {
            var deal_value = "";
            var deal_name = "";
            var company_name = "";
            var video_id = null;
            viewModel.Spinner(true);
            var sending_is_ok = true;
            if (className === "business" || className === "deal") {
                if (!!appConfig.dealValue) {
                    deal_value = eval('viewModel.' + className + '.' + appConfig.dealValue + '.value')
                } else {
                    console.log("You are missing dealValue in the config.");
                }
                deal_name = eval('viewModel.' + className + '.name.text');
                company_name = eval('viewModel.' + className + '.company.text');
            } else if (className === "company") {
                company_name = eval('viewModel.' + className + '.name.text');
            } else {
                try {
                    company_name = eval('viewModel.' + className + '.company.text');
                } catch (e) {
                    sending_is_ok = false;
                    console.log(e);
                    className = lbs.limeDataConnection.activeInspector.class.name;
                    if (className) {
                        sendDocument(automaticSending);
                    }

                }
            }

            if (sending_is_ok) {
                if (viewModel.videoData().video_id) {
                    video_id = viewModel.videoData().video_id;
                }
                if (!!viewModel.selectedTemplate() && viewModel.useTemplates()) {
                    var documentData = {
                        name: viewModel.selectedTemplate().name,
                        type: 'sales',
                        external_id: key,
                        value: deal_value,
                        recipients: [],
                        company_name: company_name,
                        is_automatic_sending: automaticSending,
                        is_reminder_sending: viewModel.smartReminders(),
                        is_sms_sending: viewModel.sendSMS(),
                        email_send_subject: viewModel.emailSubject(),
                        email_send_message: viewModel.emailMessage(),
                        video_id: video_id ? video_id : null,
                    }

                    if (viewModel.selectedTemplateFields().length > 0) {
                        var customFields = [];
                        $.each(viewModel.selectedTemplateFields(), function (index, field) {
                            customFields.push({
                                'id': field.field_id,
                                'value': field.field_value
                            });
                        });
                        documentData.custom_fields = customFields;
                    }

                    documentData.template_id = viewModel.selectedTemplateId();

                    var have_signer = viewModel.recipientsList().filter(function (i) {
                        return i.signer() == true;
                    });

                    if (have_signer != "undefiend") {
                        documentData.is_signing = true;
                    } else {
                        documentData.is_signing = 0;
                    }
                    gaRecipientList = [];
                    $.each(viewModel.recipientsList(), function (i, rec) {
                        var recipeint = new gaRecipient(rec);
                        gaRecipientList.push(recipeint);
                    });
                    documentData.recipients = gaRecipientList;
                    postDocument(documentData, automaticSending);
                } else {
                    $.each(viewModel.document(), function (i, file) {
                        var documentData = {
                            name: file.documentName(),
                            file_ids: '',
                            type: 'sales',
                            value: deal_value,
                            external_id: file.external_id,
                            recipients: [],
                            company_name: company_name,
                            is_automatic_sending: automaticSending,
                            is_reminder_sending: viewModel.smartReminders(),
                            is_sms_sending: viewModel.sendSMS(),
                            email_send_subject: viewModel.emailSubject(),
                            email_send_message: viewModel.emailMessage(),
                            video_id: video_id ? video_id : null,
                        }
                        gaRecipientList = [];
                        $.each(viewModel.recipientsList(), function (i, rec) {
                            var recipeint = new gaRecipient(rec);
                            gaRecipientList.push(recipeint);
                        });

                        //Sätter filid och filnamn till dokumentet.
                        documentData.file_ids = viewModel.uploadedDocuments().join(',');
                        documentData.recipients = gaRecipientList;

                        var have_signer = viewModel.recipientsList().filter(function (i) {
                            return i.signer() == true;
                        });

                        if (have_signer != "undefiend") {
                            documentData.is_signing = true;
                        } else {
                            documentData.is_signing = 0;
                        }

                        postDocument(documentData, automaticSending);
                    });
                }
            }
        }

        function postDocument(documentData, automaticSending) {
            apiRequest('documents', 'POST', documentData, function (data) {
                lbs.common.executeVba("GetAccept.SetDocumentStatus," + 1 + ',' + className);
                if (automaticSending) {
                    viewModel.Spinner(false);
                    lbs.common.executeVba("GetAccept.CreateHistory");
                    hideAllSteps();
                    viewModel.Reminder(true);
                } else {
                    setTimeout(function () {
                        viewModel.Spinner(false);
                        var docUrl = '/document/edit/' + data.id;
                        var sso_url = 'https://app.getaccept.com/auth/sso/login?token=' + escape(accessToken) + '&entity_id=' + entityId + '&go=' + escape(docUrl);
                        lbs.common.executeVba("GetAccept.OpenGALink", sso_url);
                        done();
                    }, 3000);
                }
                viewModel.videoData({});
                viewModel.hasVideo(false);
            });
        }

        function uploadDocument() {
            viewModel.uploadedDocuments.removeAll();
            var document_data = lbs.common.executeVba("GetAccept.GetDocumentData," + className);
            document_data = JSON.parse(document_data);
            if (document_data.length > 0) {
                $.each(document_data, function (index, doc) {
                    var file = window.atob(doc.file_content);
                    var len = file.length;
                    var bytes = new Uint8Array(len);
                    for (var i = 0; i < len; i++) {
                        bytes[i] = file.charCodeAt(i);
                    }
                    var blobData = new Blob([bytes.buffer]);
                    try {
                        var xhrRequest = new XMLHttpRequest();
                        var postUrl = apiEndpoint + "/v1/upload";
                        var formData = new FormData();
                        formData.append("file", blobData, doc.file_name);
                        xhrRequest.open('POST', postUrl, true);
                        xhrRequest.setRequestHeader('Authorization', 'bearer ' + accessToken);
                        xhrRequest.addEventListener('readystatechange', function (evt) {
                            if (xhrRequest.readyState == 4) {
                                status = xhrRequest.status;
                                if (status == 200) {
                                    viewModel.Spinner(false);
                                    var result = JSON.parse(xhrRequest.responseText);
                                    viewModel.uploadedDocuments.push(result.file_id);
                                }
                            }
                        });
                        xhrRequest.send(formData);
                    } catch (error) {
                        console.log(e);
                    }
                });
            } else {
                alert("Can't upload this document. Check so content is okey.");
                viewModel.worksteps([1, 1]);
                viewModel.File(true);
                viewModel.Recipient(false);
                viewModel.Document(false);
                viewModel.GaDocuments(false);
            }
        }

        listEntities = function () {
            viewModel.entityList.removeAll();
            apiRequest('users/me', 'GET', '', function (data) {
                viewModel.User = data.user;
                $.each(data.entities, function (i, entity) {
                    var em = new entityModel(entity, data.user.id);
                    viewModel.entityList.push(em);
                });
                viewModel.entityList().sort(function (left, right) {
                    if (right.id == entityId) {
                        return 1;
                    } else if (left.id == entityId) {
                        return -1
                    }
                });
            });
        }

        function entityModel(entityData, userHash) {
            var entity = this;
            entity.name = entityData.name;
            entity.id = entityData.id;
            entity.selectEntity = function () {
                entityId = this.id;
                apiRequest('refresh/' + entityId, 'GET', '', function (data) {
                    data.entity_id = entityId;
                    data.user_hash = userHash;
                    saveToken(data);
                    viewModel.Settings(false);
                    initGa();
                });
            }
            return entity;
        }

        function listCoworkers() {
            viewModel.coworkerList.removeAll();
            apiRequest('users', 'GET', '', function (data) {
                if (!viewModel.searchToInvite) {
                    var coworkers = lbs.common.executeVba('GetAccept.GetCoworkerList');
                    coworkers = JSON.parse(coworkers);
                    _.each(coworkers.Persons, function (coworkerData) {
                        if (coworkerData.email !== '') {
                            var existingUser = _.find(data.users, function (u) {
                                if (u.email === coworkerData.email) {
                                    return true;
                                } else {
                                    return false;
                                }
                            });

                            if (!existingUser) {
                                var c = new coworkerModel(coworkerData);
                                viewModel.coworkerList.push(c);
                            }
                        }
                    });
                }
            });

        }

        function coworkerModel(coworkerData) {
            var coworker = this;
            coworker.name = coworkerData.firstname + ' ' + coworkerData.lastname;
            coworker.email = coworkerData.email;
            coworker.sendInvite = function () {
                viewModel.coworkerList.remove(this);
                apiRequest('users', 'POST', this, function (data) {});
            }
            return coworker;
        }

        function initGa() {
            viewModel.showApp = true;
            lbs.common.executeVba("GetAccept.initGa," + appConfig.personSourceTab + ',' + appConfig.personSourceField)
            var isLoggedOn = checkLogin();
            if (isLoggedOn) {

            } else {
                //viewModel.Login(true);
            }
        }

        function setupPusher() {
            if (!pusherInit) {
                var pusher = new Pusher('d3f332f9b68a9e71641e', {
                    encrypted: true,
                    authEndpoint: apiEndpoint + '/v1/pusher',
                    disableStats: true,
                    auth: {
                        headers: {
                            'Authorization': 'bearer ' + accessToken
                        }
                    }
                });
                var pusherChannel = pusher.subscribe('private-user_' + userHash);
                pusherInit = true;
                pusherChannel.bind('video.uploaded', function (data) {
                    videoTimer = true;
                    checkVideo(data.message.job_id, data.message.video_id);
                });
            }
        }

        function checkVideo(job_id, video_id) {
            viewModel.Spinner(true);
            if (videoTimer) {
                apiRequest("video/job/" + job_id, "GET", "", function (data) {
                    if (videoTimer) {
                        if (data) {
                            if (data.Status === "Complete") {
                                videoTimer = false;
                                viewModel.hasVideo(true);
                                data.video_id = video_id;
                                viewModel.videoData(data);
                                viewModel.Spinner(false);
                            } else if (data.Status == "Error") {
                                videoTimer = false;
                            } else {
                                setTimeout(function () {
                                    checkVideo(job_id, video_id);
                                }, 2000);
                            }
                        }
                    } else {
                        setTimeout(function () {
                            checkVideo(job_id, video_id);
                        }, 2000);
                    }
                });
            }
        }

        viewModel.hidePersonList = hidePersonList;
        viewModel.showRecipientPicker = showRecipientPicker;
        viewModel.toogleGaView = toogleGaView;
        viewModel.checkLogin = checkLogin;
        viewModel.logon = logon;
        viewModel.logout = logout;
        viewModel.backToLogin = backToLogin;
        viewModel.cancel = cancel;
        viewModel.showPersonList = showPersonList;
        viewModel.showFile = showFile;
        viewModel.showDocument = showDocument;
        viewModel.send = send;
        viewModel.open = open;
        viewModel.done = done;
        viewModel.closeVideo = closeVideo;
        viewModel.uploadVideo = uploadVideo;
        viewModel.removeVideo = removeVideo;
        viewModel.backDocument = backDocument;
        viewModel.backFile = backFile;
        viewModel.backFromVideo = backFromVideo;
        viewModel.showSettings = showSettings;
        viewModel.showInvite = showInvite;
        viewModel.showHelp = showHelp;
        viewModel.createTodo = createTodo;
        viewModel.refreshGaDocuments = refreshGaDocuments;
        viewModel.showMenu = showMenu;
        viewModel.hideMenu = hideMenu;
        viewModel.createFreeAccount = createFreeAccount;
        viewModel.registerSignup = registerSignup;
        viewModel.showTemplateParameters = showTemplateParameters;
        initGa();
        return viewModel;
    };
});

function listCounties() {
    return countrylist = {
        "SE": "Sweden",
        "DK": "Denmark",
        "FI": "Finland",
        "NO": "Norway",
        "GB": "United Kingdom",
        "US": "United States",
        "OT": "Other"
    };
}