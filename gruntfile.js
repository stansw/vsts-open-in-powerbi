module.exports = function (grunt) {
    grunt.initConfig({
        ts: {
            build: {
                tsconfig: true,
            },
            watch: {
                tsconfig: true,
                watch: "."
            },
            options: {
                fast: 'never'
            }
        },
        exec: {
            package_dev: {
                command: "tfx extension create --rev-version --manifests vss-extension.json --overrides-file configs/dev.json  --output-path visx/dev",
                stdout: true,
                stderr: true
            },
            package_release: {
                command: "tfx extension create --manifests vss-extension.json --overrides-file configs/release.json --output-path visx/release",
                stdout: true,
                stderr: true
            },
            publish_dev: {
                command: "tfx extension publish --service-url https://marketplace.visualstudio.com --manifests vss-extension.json --overrides-file configs/dev.json",
                stdout: true,
                stderr: true
            },
            publish_release: {
                command: "tfx extension publish --service-url https://marketplace.visualstudio.com --manifests vss-extension.json --overrides-file configs/release.json",
                stdout: true,
                stderr: true
            }
        },
        copy: {
            scripts: {
                files: [
                    {
                        expand: true,
                        flatten: true,
                        src: ["node_modules/vss-web-extension-sdk/lib/VSS.SDK.min.js"],
                        dest: "out/lib/"
                    },
                    {
                        expand: true,
                        flatten: true,
                        src: ["node_modules/telemetryclient-team-services-extension/lib/telemetryclient.js"],
                        dest: "out/lib/"
                    },
                    {
                        expand: true,
                        flatten: true,
                        src: ["node_modules/applicationinsights-js/dist/ai.0.js"],
                        dest: "out/lib/"
                    },
                    {
                        expand: true,
                        flatten: true,
                        src: ["node_modules/jszip/dist/jszip.min.js"],
                        dest: "out/lib/"
                    },
                    {
                        expand: true,
                        flatten: true,
                        src: ["lib/jquery.binarytransport.js"],
                        dest: "out/lib/"
                    },
                    {
                        expand: true,
                        flatten: true,
                        src: ["lib/FileSaver.min.js"],
                        dest: "out/lib/"
                    }
                ]
            }
        },

        clean: ["scripts/**/*.js", "*.vsix", "build", "test"],

        karma: {
            unit: {
                configFile: 'karma.conf.js',
                singleRun: true,
                browsers: ["PhantomJS"]
            }
        }
    });

    grunt.loadNpmTasks("grunt-ts");
    grunt.loadNpmTasks("grunt-exec");
    grunt.loadNpmTasks("grunt-contrib-copy");
    grunt.loadNpmTasks('grunt-contrib-clean');
    grunt.loadNpmTasks('grunt-karma');

    grunt.registerTask("build", ["ts:build", "copy:scripts"]);

    grunt.registerTask("test", ["ts:buildTest", "karma:unit"]);

    grunt.registerTask("package-dev", ["build", "exec:package_dev"]);
    grunt.registerTask("package-release", ["build", "exec:package_release"]);
    grunt.registerTask("publish-dev", ["package-dev", "exec:publish_dev"]);
    grunt.registerTask("publish-release", ["package-release", "exec:publish_release"]);

    grunt.registerTask("watch-build", ["copy:scripts", "ts:watch"]);

    grunt.registerTask("default", ["package-dev"]);
};