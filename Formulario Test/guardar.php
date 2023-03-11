<?php

// Autenticación con la API de Microsoft Graph
$token_url = "https://login.microsoftonline.com/{your_tenant_id}/oauth2/v2.0/token";
$client_id = "{your_client_id}";
$client_secret = "{your_client_secret}";
$resource = "https://graph.microsoft.com/";
$grant_type = "client_credentials";

$data = array(
    "grant_type" => $grant_type,
    "client_id" => $client_id,
    "client_secret" => $client_secret,
    "resource" => $resource
);

$options = array(
    "http" => array(
        "header" => "Content-type: application/x-www-form-urlencoded\r\n",
        "method" => "POST",
        "content" => http_build_query($data)
    )
);

$context = stream_context_create($options);
$token_response = file_get_contents($token_url, false, $context);
$token = json_decode($token_response)->access_token;

// Creación del objeto de solicitud para la API de Microsoft Graph
$url = "https://graph.microsoft.com/v1.0/drives/{your_drive_id}/items/{your_folder_id}/workbook/tables('Sheet1')/rows";
$header = array(
    "Authorization: Bearer $token",
    "Content-Type: application/json"
);

// Recopilación de datos del formulario
$nombre = $_POST['nombre'];
$telefono = $_POST['telefono'];
$email = $_POST['email'];
$id = $_POST['id'];

// Creación del objeto de datos para enviar a la hoja de cálculo
$data = array(
    array(
        "values" => array(
            $nombre,
            $telefono,
