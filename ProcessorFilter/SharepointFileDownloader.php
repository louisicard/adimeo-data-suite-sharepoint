<?php

namespace AdimeoDataSuite\ProcessorFilter;

use AdimeoDataSuite\Model\Datasource;
use AdimeoDataSuite\Model\ProcessorFilter;
use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\Runtime\Utilities\RequestOptions;
use Office365\PHP\Client\SharePoint\ClientContext;

class SharepointFileDownloader extends ProcessorFilter
{
  function getDisplayName()
  {
    return "Sharepoint file downloader";
  }

  function getSettingFields()
  {
    return array(
      'company_url' => array(
        'type' => 'text',
        'label' => 'Company url (E.g.: https://mycompany.sharepoint.com)',
        'required' => true
      ),
      'username' => array(
        'type' => 'text',
        'label' => 'Username',
        'required' => true
      ),
      'password' => array(
        'type' => 'text',
        'label' => 'Password',
        'required' => true
      )
    );
  }

  function getFields()
  {
    return array('filePath');
  }

  function getArguments()
  {
    array(
      'siteName' => 'Site name',
      'relativePath' => 'Relative path'
    );
  }

  function execute(&$document, Datasource $datasource)
  {
    $authCtx = new AuthenticationContext($this->getSettings()['company_url']);
    $authCtx->acquireTokenForUser($this->getSettings()['username'], $this->getSettings()['password']);

    $downloadUrl = $this->getArgumentValue('siteName', $document) . "/_api/web/GetFileByServerRelativeUrl('" . rawurlencode($this->getArgumentValue('relativePath', $document)) . "')/\$value?@target='" . urlencode($this->getSettings()['company_url']) . "'";
    $fileRequest = new RequestOptions($downloadUrl);
    $ctxFile = new ClientContext($downloadUrl, $authCtx);
    $content = $ctxFile->executeQueryDirect($fileRequest);
    $tempFile = tempnam(sys_get_temp_dir(), 'ads_sp_');
    $datasource->getOutputManager()->writeLn('>>> Downloading file ' . $this->getArgumentValue('relativePath', $document));
    file_put_contents($tempFile, $content);

    return array(
      'filePath' => $tempFile
    );
  }

}