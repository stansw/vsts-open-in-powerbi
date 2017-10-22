module.exports = function (grunt) {
    grunt.initConfig({
        ts: {
            build: {
                tsconfig: true,
                outDir: "./out/release",
                src: ["./src/**/*.ts"],
                options: {
                    module: "amd",
                    sourceMap: false
                }
            },
            buildDebug: {
                tsconfig: true,
                outDir: "./out/debug",
                watch: ".",
                options: {
                    module: "umd",
                    sourceMap: true
                },
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

        clean: ["out/**/*.js", "out/**/*.js.map", "*.vsix"],
    });

    grunt.loadNpmTasks("grunt-ts");
    grunt.loadNpmTasks("grunt-exec");
    grunt.loadNpmTasks("grunt-contrib-copy");
    grunt.loadNpmTasks('grunt-contrib-clean');

    grunt.registerTask("build", ["ts:build"]);

    grunt.registerTask("debug", ["clean", "ts:buildDebug"]);

    grunt.registerTask("package-dev", ["build", "exec:package_dev"]);
    grunt.registerTask("package-release", ["build", "exec:package_release"]);
    grunt.registerTask("publish-dev", ["package-dev", "exec:publish_dev"]);
    grunt.registerTask("publish-release", ["package-release", "exec:publish_release"]);

    grunt.registerTask("default", ["package-dev"]);
};