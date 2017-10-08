<!doctype html>
<html lang="{{ app()->getLocale() }}">
    <head>
        <meta charset="utf-8">
        <meta http-equiv="X-UA-Compatible" content="IE=edge">
        <meta name="viewport" content="width=device-width, initial-scale=1">

        <title>Laravel</title>

        <!-- Fonts -->
        <link href="https://fonts.googleapis.com/css?family=Raleway:100,600" rel="stylesheet" type="text/css">

        <!-- Styles -->
        <style>
            html, body {
                background-color: #fff;
                color: #636b6f;
                font-family: 'Raleway', sans-serif;
                font-weight: 100;
                height: 100vh;
                margin: 0;
            }

            .full-height {
                height: 100vh;
            }

            .flex-center {
                align-items: center;
                display: flex;
                justify-content: center;
            }

            .position-ref {
                position: relative;
            }

            .top-right {
                position: absolute;
                right: 10px;
                top: 18px;
            }

            .content {
                text-align: center;
            }

            .title {
                font-size: 84px;
            }

            .links > a {
                color: #636b6f;
                padding: 0 25px;
                font-size: 12px;
                font-weight: 600;
                letter-spacing: .1rem;
                text-decoration: none;
                text-transform: uppercase;
            }

            .m-b-md {
                margin-bottom: 30px;
            }
        </style>
        <script src="js/jquery-3.2.1.min.js"></script>
        <script src="js/docxtemplater.js"></script>
        <script src="js/docxtemplater-image-module.js"></script>
        <script src="js/jszip.js"></script>
        <script src="js/file-saver.min.js"></script>
        <script src="js/jszip-utils.js"></script>
        <script type="text/javascript">
            $(function() {
                $(".generateDocument").click(function() {
                    $.ajax({
                        type: "GET",
                        url: "{{route('generateDocument')}}",
                        success: function(response) {
                            var width = response[14];
                            var height = response[15];
                            loadFile(response[13], function (error, content) {
                                var myimg = content;
                                
                                var opts = {};
                                opts.centered = false;
                                opts.getImage = function (tagValue, tagName) {
                                  return myimg;
                                };

                                opts.getSize = function (img, tagValue, tagName) {
                                  return [width/(height/60), 60];
                                };
                                var imageModule = new window.ImageModule(opts);

                                loadFile("template.docx",function(error,content){
                                    if (error) { throw error };
                                    
                                    var zip = new JSZip(content);
                                    var doc = new Docxtemplater().loadZip(zip);
                                    doc.attachModule(imageModule);
                                    doc.setData({
                                        document_number: response[1],
                                        document_title: response[2],
                                        logo: response[13],
                                        product_name: response[4],
                                        product_no: response[5],
                                        risk_mgmt_standard: response[6],
                                        document_purpose: response[7],
                                        risk_mgmt_activities: response[8],
                                        prepared_by: response[9],
                                        approved_by: response[11],
                                        designation_preparer: response[10],
                                        designation_approver: response[12],
                                        version: response[3]
                                    });

                                    try {
                                        // render the document (replace all occurences of {first_name} by John, {last_name} by Doe, ...)
                                        doc.render();
                                    }
                                    catch (error) {
                                        var e = {
                                            message: error.message,
                                            name: error.name,
                                            stack: error.stack,
                                            properties: error.properties,
                                        };
                                        console.log(JSON.stringify({error: e}));
                                        // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
                                        throw error;
                                    }

                                    var out=doc.getZip().generate({
                                        type:"blob"
                                    }); //Output the document using Data-URI
                                    saveAs(out,response[0] + ".docx");
                                });
                            });
                        }
                    });
                    
                    function loadFile(url, callback)
                    {
                        JSZipUtils.getBinaryContent(url, callback);
                    }
                });
            });
        </script>
    </head>
    <body>
        <div class="flex-center position-ref full-height">
            @if (Route::has('login'))
                <div class="top-right links">
                    @auth
                        <a href="{{ url('/home') }}">Home</a>
                    @else
                        <a href="{{ route('login') }}">Login</a>
                        <a href="{{ route('register') }}">Register</a>
                    @endauth
                </div>
            @endif

            <div class="content">
                <div class="title m-b-md">
                    Laravel
                </div>

                <div class="links">
                    <a href="#" class="generateDocument">Generate Document</a>
                    <a href="https://laravel.com/docs">Documentation</a>
                    <a href="https://laracasts.com">Laracasts</a>
                    <a href="https://laravel-news.com">News</a>
                    <a href="https://forge.laravel.com">Forge</a>
                    <a href="https://github.com/laravel/laravel">GitHub</a>
                </div>
            </div>
        </div>
    </body>
</html>
