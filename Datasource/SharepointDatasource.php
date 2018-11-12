<?php

namespace AdimeoDataSuite\Datasource;

use AdimeoDataSuite\Exception\DatasourceExecutionException;
use AdimeoDataSuite\Model\Datasource;
use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\Runtime\Utilities\RequestOptions;
use Office365\PHP\Client\SharePoint\ClientContext;

class SharepointDatasource extends Datasource
{

  const SHAREPOINT_PAGER_SIZE = 500;

  function getOutputFields()
  {
    return array(
      'docPath',
      'relativePath',
      'siteName',
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
      'last_modified_time' => array(
        'type' => 'text',
        'label' => 'Last modified date (YYYY-MM-dd HH:mm:ss)',
        'required' => true
      )
    );
  }

  function getDisplayName()
  {
    return 'Sharepoint datasource';
  }

  function execute($args)
  {
    if(isset($args['last_modified_time'])) {
      $lastModifiedTime = \DateTime::createFromFormat('Y-m-d H:i:s', $args['last_modified_time']);
    }
    if(isset($lastModifiedTime) && $lastModifiedTime) {
      $this->querySharepoint($lastModifiedTime);
    }
    else {
      throw new DatasourceExecutionException('Argument last_modified_time is not a valid date/time (format should be YYYY-MM-dd HH:mm:ss)');
    }
  }

  private function querySharepoint(\DateTime $lastModifiedTime, $from = 0) {
    $authCtx = new AuthenticationContext($this->getSettings()['company_url']);
    $authCtx->acquireTokenForUser($this->getSettings()['username'], $this->getSettings()['password']);

    $selectProperties = ['Path', 'LastModifiedTime', 'SiteName'];
    if(isset($this->getSettings()['select_properties']) && !empty($this->getSettings()['select_properties'])) {
      foreach(array_map('trim', explode(',', $this->getSettings()['select_properties'])) as $prop) {
        if(!in_array($prop, $selectProperties)) {
          $selectProperties[] = $prop;
        }
      }
    }

    $searchQuery = "'(" . $this->getSettings()['search_request'] . ") AND LastModifiedTime>" . $lastModifiedTime->format('Y-m-d\TH:i:s') . " AND IsDocument:true'";

    $searchUrl = trim($this->getSettings()['company_url'], '/')
      . "/_api/search/query?"
      . "querytext=" . rawurlencode($searchQuery)
      . "&selectproperties=" . rawurlencode("'" . implode(',', $selectProperties) . "'")
      . "&sortlist=" . rawurlencode("'LastModifiedTime:descending'")
      . "&rowlimit=" . static::SHAREPOINT_PAGER_SIZE
      . "&startrow=" . $from;

    $request = new RequestOptions($searchUrl);
    $ctx = new ClientContext($searchUrl, $authCtx);
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
        'docPath' => isset($doc['Path']) ? $doc['Path'] : null,
        'relativePath' => null,
        'siteName' => isset($doc['SiteName']) ? $doc['SiteName'] : null,
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
    }
    if($count > 0) {
      $this->querySharepoint($lastModifiedTime, $from + static::SHAREPOINT_PAGER_SIZE);
    }

  }

}