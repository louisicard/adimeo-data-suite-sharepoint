<?php

namespace AdimeoDataSuite\Datasource;

use AdimeoDataSuite\Exception\DatasourceExecutionException;
use AdimeoDataSuite\Model\Datasource;
use Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use Office365\PHP\Client\Runtime\Utilities\RequestOptions;
use Office365\PHP\Client\SharePoint\Change;
use Office365\PHP\Client\SharePoint\ChangeQuery;
use Office365\PHP\Client\SharePoint\ChangeType;
use Office365\PHP\Client\SharePoint\ClientContext;
use Office365\PHP\Client\SharePoint\SPList;

class SharepointChangeLogs extends Datasource
{

  const SHAREPOINT_PAGER_SIZE = 500;

  function getOutputFields()
  {
    return array(
      'authContext',
      'operation',
      'uniqueId',
      'siteName'
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
    return 'Sharepoint change logs';
  }

  function execute($args)
  {
    $lastModified = \DateTime::createFromFormat('Y-m-d H:i:s', $args['last_modified_time']);
    if(!$lastModified) {
      throw new DatasourceExecutionException('Argument last_modified_time is incorrect. Expected format is Y-m-d H:i:s');
    }

    if($this->authContext == null) {
      $this->authContext = new AuthenticationContext($this->getSettings()['company_url']);
      $this->authContext->acquireTokenForUser($this->getSettings()['username'], $this->getSettings()['password']);
    }

    //Get Sites
    $this->getOutputManager()->writeLn('Searching for sites...');
    $sites = [];
    $this->searchSites($sites);
    //Exclude personal sites
    foreach($sites as $i => $site) {
      $site_r = explode('/', $site);
      if($site_r[count($site_r) - 2] == 'personal') {
        unset($sites[$i]);
      }
    }
    $sites = array_values($sites);

    //Getting changes for every site
    foreach($sites as $site) {
      $this->getOutputManager()->writeLn('Getting logs for site ' . $site);
      $ctx = new ClientContext($site, $this->authContext);
      $list = $ctx->getWeb()->getLists()->getByTitle('Documents');
      $logs = ['to_index' => [], 'to_delete' => []];
      try {
        $this->getChanges($this->authContext, $list, $lastModified, $logs);
      }
      catch(\Exception $ex) {
        //Nothing to do
      }
      foreach($logs as $op => $entries) {
        foreach($entries as $uniqueId => $changeToken) {
          $this->index(array(
            'authContext' => $this->authContext,
            'operation' => $op,
            'uniqueId' => $uniqueId,
            'siteName' => strtolower($site)
          ));
        }
      }
      $this->getOutputManager()->writeLn(count($logs['to_index']) . ' documents to index');
      $this->getOutputManager()->writeLn(count($logs['to_delete']) . ' documents to delete');
    }
  }

  /**
   * @var AuthenticationContext
   */
  private $authContext = null;

  private function getChanges(AuthenticationContext $authCtx, SPList $list, \DateTime $lastModified, &$logs, $token = null) {
    $ctx = $list->getContext();
    $query = new ChangeQuery();
    $query->Add = true;
    $query->Update = true;
    $query->DeleteObject = true;
    $query->Item = true;
    $query->File = true;
    if($token != null) {
      $query->ChangeTokenStart = $token;
    }

    $changes = $list->getChanges($query);
    $ctx->executeQuery();

    $lastToken = null;
    foreach ($changes->getData() as $change) {
      /** @var Change $change */
      $changeTypeName = ChangeType::getName($change->ChangeType);
      $lastToken = $change->ChangeToken;
      $time = \DateTime::createFromFormat('Y-m-d\TH:i:s\Z', $change->Time);
      $properties = $change->getProperties();
      if($time && isset($properties['UniqueId'])) {
        if(!$lastModified->diff($time)->invert) {
          if ($changeTypeName == 'Add' || $changeTypeName == 'Update') {
            $logs['to_index'][$properties['UniqueId']] = serialize($change->ChangeToken);
          } elseif ($changeTypeName == 'DeleteObject') {
            $logs['to_delete'][$properties['UniqueId']] = serialize($change->ChangeToken);
            if (in_array($properties['UniqueId'], array_keys($logs['to_index']))) {
              unset($logs['to_index'][$properties['UniqueId']]);
            }
          }
        }
      }
    }
    if($lastToken != null) {
      $this->getChanges($authCtx, $list, $lastModified, $logs, $lastToken);
    }
  }

  private function searchSites(&$sites, $from = 0) {
    $searchQuery = "'contentclass:STS_Site OR contentclass:STS_Web'";

    $searchUrl = trim($this->getSettings()['company_url'], '/')
      . "/_api/search/query?"
      . "querytext=" . rawurlencode($searchQuery)
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
      $count++;
      $sites[] = $doc['Path'];
    }
    if($count > 0) {
      $this->searchSites($sites, $from + static::SHAREPOINT_PAGER_SIZE);
    }

  }

}