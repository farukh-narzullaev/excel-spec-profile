<?php

class CustomException extends RuntimeException {

}

function execute() {
    throw new Exception('Exception Executed.');
}

try {
    execute();
} catch (CustomException $e) {
    echo $e->getMessage().PHP_EOL;
}

