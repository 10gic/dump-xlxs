#/bin/bash

# dump xlsx to text-based table

printUsage() {
    echo "Dump xlxs file to text-based table" 1>&2
    echo "Usage 1: dumpXlxs.sh [-d logLevel] [-t type] input.xlsx [output.txt]" 1>&2;
    echo "Usage 2: dumpXlxs.sh [-d logLevel] -r input.txt output.xlsx" 1>&2;
    echo "" 1>&2;
    echo "-d logLevel, specify log level" 1>&2;
    echo "   logLevel = trace|debug|info|warn|error, default is info" 1>&2;
    echo "-t type, specify type of text-based table"  1>&2;
    echo "   type = emacs|orgmode, default is emacs" 1>&2;
    echo "-r (means reverse), generate emacs table back into xlxs file" 1>&2;
}

logLevel="info"
type="emacs"
reverse=0

while getopts "hd:t:r" opt; do
    case $opt in
        h ) printUsage
            exit
            ;;
        d ) logLevel=$OPTARG
            ;;
        t ) type=$OPTARG
            ;;
        r)
            reverse=1
            ;;
        \? ) printUsage
             exit
             ;;
    esac
done
shift $(($OPTIND - 1))

if [[ $reverse -eq 0 ]]; then
    if [[ $type == "emacs" ]]; then
        java -Dorg.slf4j.simpleLogger.defaultLogLevel=${logLevel:-"info"} -jar target/dumpXlsx-1.0.jar "$@"
    elif
        [[ $type == "orgmode" ]]; then
        java -Dorg.slf4j.simpleLogger.defaultLogLevel=${logLevel:-"info"} -cp target/dumpXlsx-1.0.jar com.aandds.app.Xlsx2OT "$@"
    else
        printUsage
        exit
    fi
else
    java -Dorg.slf4j.simpleLogger.defaultLogLevel=${logLevel:-"info"} -cp target/dumpXlsx-1.0.jar com.aandds.app.ET2Xlsx "$@"
fi
