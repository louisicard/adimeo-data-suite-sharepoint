<?php

namespace AdimeoDataSuite\ProcessorFilter;

use AdimeoDataSuite\Model\Datasource;
use AdimeoDataSuite\Model\ProcessorFilter;
use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\Runtime\Utilities\RequestOptions;
use Office365\PHP\Client\SharePoint\ClientContext;

class SharepointDocumentRetriever extends ProcessorFilter
{
  function getDisplayName()
  {
    return "Sharepoint document retriever";
  }

  function getSettingFields()
  {
    return array(
      'docs_only' => array(
        'type' => 'boolean',
        'label' => 'Retrieve documents only (Stops processor otherwise)',
        'required' => false
      )
    );
  }

  function getFields()
  {
    return array('doc');
  }

  function getArguments()
  {
    return array(
      'authContext' => 'Authentication context',
      'siteName' => 'Site name',
      'itemId' => 'Item ID (in site document library)'
    );
  }

  function execute(&$document, Datasource $datasource)
  {
    /** @var AuthenticationContext $authCtx */
    $authCtx = $this->getArgumentValue('authContext', $document);

    $url = $this->getArgumentValue('siteName', $document) . "/_api/web/lists/getByTitle('Documents')/items?\$select=EncodedAbsUrl,FileSystemObjectType&\$filter=" . rawurlencode('Id eq ' . $this->getArgumentValue('itemId', $document));
    $request = new \Office365\PHP\Client\Runtime\Utilities\RequestOptions($url);
    $ctx = new ClientContext($url, $authCtx);
    $data = $ctx->executeQueryDirect($request);

    $data = json_decode($data, TRUE);
    if(isset($data['d']['results'][0])) {
      $props = $data['d']['results'][0];
      if($this->getSettings()['docs_only'] && $props['FileSystemObjectType'] === 0 || !$this->getSettings()['docs_only']) {
        $path = $props['EncodedAbsUrl'];
        return array(
          'doc' => $this->searchForDocument($authCtx, $path)
        );
      }
      else {
        $document = [];
        return NULL;
      }
    }
    return array(
      'doc' => NULL
    );
  }

  private function searchForDocument(AuthenticationContext $authContext, $path) {
    $searchQuery = "'Path:\"" . $path . "\"'";

    $path_r = explode('//', $path);
    $companyUrl = 'https://' . explode('/', $path_r[1])[0];

    $searchUrl = trim($companyUrl, '/')
      . "/_api/search/query?"
      . "querytext=" . rawurlencode($searchQuery);

    $request = new RequestOptions($searchUrl);
    $ctx = new ClientContext($searchUrl, $authContext);
    $data = $ctx->executeQueryDirect($request);

    $data = json_decode($data, TRUE);
    if(isset($data['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results'][0])) {
      $doc = [];
      $row = $data['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results'][0];
      foreach ($row['Cells']['results'] as $cell) {
        if (isset($cell['Key']) && isset($cell['Value'])) {
          $doc[$cell['Key']] = $cell['Value'];
        }
      }
      return $doc;
    }

    return null;
  }

}