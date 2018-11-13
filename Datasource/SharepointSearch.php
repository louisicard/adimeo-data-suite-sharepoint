<?php

namespace AdimeoDataSuite\Datasource;

use AdimeoDataSuite\Model\Datasource;
use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\Runtime\Utilities\RequestOptions;
use Office365\PHP\Client\SharePoint\ClientContext;

class SharepointSearch extends Datasource
{

  const SHAREPOINT_PAGER_SIZE = 500;

  function getOutputFields()
  {
    return array(
      'authContext',
      'docId',
      'docPath',
      'relativePath',
      'siteName',
      'uniqueId',
      'properties'
    );
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
      ),
      'search_request' => array(
        'type' => 'text',
        'label' => 'Sharepoint search request (See doc https://docs.microsoft.com/fr-fr/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference)',
        'required' => true
      ),
      'select_properties' => array(
        'type' => 'text',
        'label' => 'Select properties (comma separated)',
        'required' => false
      )
    );
  }

  function getExecutionArgumentFields()
  {
    return array(
      'search_request' => array(
        'type' => 'text',
        'label' => 'Sharepoint search request (See doc https://docs.microsoft.com/fr-fr/sharepoint/dev/general-development/keyword-query-language-kql-syntax-reference)',
        'required' => true,
        'default_from_settings' => true
      )
    );
  }

  function getDisplayName()
  {
    return 'Sharepoint search';
  }

  function execute($args)
  {
    $this->querySharepoint($args['search_request']);
    $this->getOutputManager()->writeLn('Found ' . $this->globalCount . ' documents');
  }

  private $authContext = null;
  private $globalCount = 0;

  private function querySharepoint($searchRequest, $from = 0) {
    if($this->authContext == null) {
      $this->authContext = new AuthenticationContext($this->getSettings()['company_url']);
      $this->authContext->acquireTokenForUser($this->getSettings()['username'], $this->getSettings()['password']);
    }

    $selectProperties = ['Path', 'LastModifiedTime', 'SiteName', 'UniqueId'];
    if(isset($this->getSettings()['select_properties']) && !empty($this->getSettings()['select_properties'])) {
      foreach(array_map('trim', explode(',', $this->getSettings()['select_properties'])) as $prop) {
        if(!in_array($prop, $selectProperties)) {
          $selectProperties[] = $prop;
        }
      }
    }

    $searchQuery = "'(" . $searchRequest . ") AND IsDocument:true'";

    $searchUrl = trim($this->getSettings()['company_url'], '/')
      . "/_api/search/query?"
      . "querytext=" . rawurlencode($searchQuery)
      . "&selectproperties=" . rawurlencode("'" . implode(',', $selectProperties) . "'")
      . "&sortlist=" . rawurlencode("'LastModifiedTime:descending'")
      . "&rowlimit=" . static::SHAREPOINT_PAGER_SIZE
      . "&startrow=" . $from;

    $request = new RequestOptions($searchUrl);
    $ctx = new ClientContext($searchUrl, $this->authContext);
    $data = $ctx->executeQueryDirect($request);

    $data = json_decode($data, TRUE);
    $count = 0;
    foreach($data['d']['query']['PrimaryQueryResult']['RelevantResults']['Table']['Rows']['results'] as $row) {
      $doc = [];
      foreach($row['Cells']['results'] as $cell) {
        if(isset($cell['Key']) && isset($cell['Value'])) {
          $doc[$cell['Key']] = $cell['Value'];
        }
      }
      $to_index = [
        'authContext' => $this->authContext,
        'docId' => isset($doc['DocId']) ? $doc['DocId'] : null,
        'docPath' => isset($doc['Path']) ? $doc['Path'] : null,
        'relativePath' => null,
        'siteName' => isset($doc['SiteName']) ? $doc['SiteName'] : null,
        'uniqueId' => isset($doc['UniqueId']) ? trim($doc['UniqueId'], '{}') : null,
        'properties' => $doc
      ];
      if(isset($doc['Path'])) {
        $tmp_r = explode('//', $doc['Path']);
        $tmp_rr = explode('/', $tmp_r[1]);
        $tmp_rr = array_slice($tmp_rr, 1);
        $relativePath = '/' . implode('/', $tmp_rr);
        $to_index['relativePath'] = $relativePath;
      }
      $this->index($to_index);
      $count++;
      $this->globalCount++;
    }
    if($count > 0) {
      $this->querySharepoint($searchRequest, $from + static::SHAREPOINT_PAGER_SIZE);
    }

  }

}