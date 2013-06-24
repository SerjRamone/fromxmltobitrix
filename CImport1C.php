<?php

/**
 * Class CImport1C
 * Abstract class for import data from 1C
 * @author Sergey Greznov <segey.greznov@gmail.com>
 *
 */
abstract class CImport1C {

	/**
	 * Start memory usage
	 * @var int Start import memory size for logging
	 */
	protected $startMemoryUsage;

	/**
	 * Total memory usage
	 * @var int Total memory usage in a end of import
	 */
	protected $totalMemoryUsage;

	/**
	 * Total time
	 * @var int Total time in a end of import
	 */
	protected $totalTime;

	/**
	 * Start import time
	 * @var int Start import time for logging
	 */
	protected $timeStart;

	/**
	 * Summary log array about import
	 * @var array Summary log
	 */
	protected $log = array();

	/**
	 * Array of the admins mails for
	 * sending report about import status
	 * @var array Emails for mailing
	 */
	protected $adminsMails = array('sergey.greznov@gmail.com');

	/**
	 * Count of import errors
	 * @var int Count of errors
	 */
	protected $errorCnt = 0;

	/**
	 * Count of XML objects after XML parsing into array
	 * @var int Count of XML objects
	 */
	protected $countXMLObjects = 0;

	/**
	 * IBlock ID with departments
	 * @const int
	 */
	const depBlockID = 5;

	/**
	 * Array of users where key is a user login
	 * @var array By login=>id
	 */
	protected $arUsersByLogin = array();

	/**
	 * Array of users where key is a user XML_ID
	 * @var array By XML_ID=>id
	 */
	protected $arUsersByXmlId = array();

	/**
	 * Array of users where key is a user XML_ID
	 * @var array By NAME=>id
	 */
	protected $arUsersByName = array();

	/**
	 * Array with objects from XML after parsing
	 * @var array Array after parsing
	 */
	protected $arXMLObjects = array();

	/**
	 * Path in file system to XML file
	 * @var string Path to XML
	 */
	protected $pathToXML = '';

	/**
	 * Message text for admin mail
	 * @var string Message text
	 */
	public $messageBody = '';

	/**
	 * Object constructor
	 * @param $params array with init objects parameters
	 */
	public function __construct ($params = array()) {
		$this->startMemoryUsage = memory_get_usage();
		$mtime = microtime();
		$mtime = explode(' ', $mtime);
		$this->timeStart = $mtime[1] + $mtime[0];
		$this->pathToXML = $params['PATH_TO_XML'];
		try {
			if (!file_exists($this->pathToXML)) {
				throw new Exception('XML not exists');
			}
		} catch (Exception $e) {
			echo '<h4>'.$e->getMessage().'</h4>';
		}

		$this->arXMLObjects = $this->parseXML();
		$this->countXMLObjects = count($this->arXMLObjects);
	}

	/**
	 * Putting summary log about import
	 * @param $type string Tpype of import
	 * @return int|bool ID new row of log table or false if failure
	 */
	public function putLog ($type) {
		$arExpl = explode('/', $this->pathToXML);
		$filename = $arExpl[count($arExpl)-1];

		$mtime = microtime();
		$mtime = explode(' ',$mtime);
		$mtime = $mtime[1] + $mtime[0];
		$this->log['TOTAL_TIME'] = ($mtime - $this->timeStart);
		$end_memory_usage = memory_get_usage();
		$this->log['TOTAL_MEMORY_USAGE'] = $end_memory_usage - $this->startMemoryUsage;

		$arImportLogFields = array(
			'DATE' 			=> "'".date('Y-m-d H:i:s')."'",
			'FILE_NAME'		=> "'".$filename."'",
			'LOG'			=> "'".serialize($this->log)."'",
			'IMPORT_TYPE'	=> "'".$type."'",
			'ERROR_COUNT'	=> "'".$this->errorCnt."'",
		);

		GLOBAL $DB;

		return $DB->Insert('UT_IMPORT1C', $arImportLogFields, '', false, '', true);
	}

	/**
	 * Sending message to admins
	 * @param $subj string Subject of message for admins
	 * @param $message string Admin's message body
	 * @return void
	 */
	public function sendAdminMessage ($subj, $message) {
		foreach ($this->adminsMails as $mail) {
			bxmail($mail, $subj, $message);
		}
	}

	/**
	 * Parse XML multiline into a array
	 * @return array Array with XML elements
	 */
	protected function parseXML () {
		$fo = fopen($this->pathToXML, 'r');
		$xmlArray = array();
		$user = '';
		$putting = false;

		while ($line = fgets($fo)) {
			if (preg_match('/<\/ФизическоеЛицо>/', $line)) {
				$tmp = $this->objToArray(simplexml_load_string($user.$line));
				$xmlArray[] = $tmp;
				$user = '';
				$putting = false;
			}
			if (preg_match('/<ФизическоеЛицо>/', $line)) {
				$putting = true;
			}
			if ($putting) {
				$user .= $line;
			}
		}
		return $xmlArray;
	}

	/**
	 * Recursive convert object into array
	 * @param $o Object will be converted into array
	 * @return array
	 */
	protected function objToArray ($o) {
		$o = get_object_vars($o);
		foreach ($o as $k=>$v) {
			if(is_object($v)) {
				$o[$k] = $this->objToArray($v);
				if(empty($o[$k])) {
					$o[$k] = '';
				}
			}
		}
		return $o;
	}

	protected function selectUsers () {
		$rsUsers = CUser::GetList($by, $order, array());
		while ($arU = $rsUsers->Fetch()) {
			$this->arUsersByLogin[mb_strtoupper($arU['LOGIN'])]  = $arU['ID'];
			if (!empty($arU['XML_ID'])) {
				$this->arUsersByXmlId[$arU['XML_ID']] = $arU['ID'];
			}
			//----------------------------------------------
			if (!empty($arU['NAME'])) {
				list($name, $secondName) = explode(' ',$arU['NAME']);
				$arU['NAME'] = $name;
				if (!empty($secondName)) {
					$arU['SECOND_NAME'] = $secondName;
				}
				$key = mb_strtoupper($arU['NAME'].((!empty($arU['SECOND_NAME']))? '_'.$arU['SECOND_NAME']:'').'_'.$arU['LAST_NAME'].'|'.$arU['ID'], 'UTF-8'); //ИВАН_ИВАНОВИЧ_ИВАНОВ|456 || ИВАН_ИВАНОВ|456
				$this->arUsersByName[$key] = $arU['ID'];
			}
		}
	}

	/**
	 * Do Users import
	 * @return int|bool ID new row in log table or false if failure
	 */
	abstract public function doImport();

}


/**
 * Class CUserImport1C
 * Import users from 1C
 * @author Sergey Greznov <segey.greznov@gmail.com>
 */
class CUserImport1C extends CImport1C {

	/**
	 * Array with departments
	 * @var array Key - XML_ID, value - section ID
	 */
	private $arDepartments = array();

	/**
	 * Default password
	 * @var string Default password
	 */
	public $defaultPass = '1111111';

	/**
	 * Select departments from Intranet
	 * @return void
	 */
	private function selectDepartments () {
		$arFilter = array('IBLOCK_ID'=>self::depBlockID, 'GLOBAL_ACTIVE'=>'Y');
		$rsSect = CIBlockSection::GetList(array($by=>$order), $arFilter, false);

		while($arSect = $rsSect->Fetch()) {
			if (strlen(trim($arSect['XML_ID'])) > 0) {
				$this->arDepartments[trim($arSect['XML_ID'])] = $arSect['ID'];
			}
		}
	}

	/**
	 * Do Users import
	 * @return int|bool ID new row in log table or false if failure
	 */
	public function doImport() {
		try {
			if (!CModule::IncludeModule('iblock') || !CModule::IncludeModule('main')) {
				throw new Exception('Bitrix modules "iblock" or "main" not included');
			}
		} catch (Exception $e) {
			echo '<h4>'.$e->getMessage().'</h4>';
		}

		$this->selectDepartments();
		$this->selectUsers();

		$this->log = array (
			'PROC_ROWS'			=> 0,
			'USERS_UPDATED'		=> 0,
			'USERS_ADDED'		=> 0,
			'USERS_WITHOUT_DEP'	=> 0,
			'ADDING_ERRORS'		=> array(),
			'UPDATING_ERRORS'	=> array(),
			'COLLIZIONS'		=> array()
		);

		//******************************************
		/*g($this->arDepartments);
		g($this->arUsersByLogin);
		g($this->arUsersByXmlId);
		g($this->arUsersByName);
		die('xxx');*/
		//******************************************
		#var_dump($this->countXMLObjects);die;

		if ($this->countXMLObjects == 0) {
			$this->errorCnt = 1;
			$this->messageBody .= 'Пустой (или не валидный) файл импорта.';
		}

		for ($i = 0; $i < $this->countXMLObjects; $i++) {

			$u = $this->arXMLObjects[$i];
			if (!array_key_exists($u['ИДПодразделение'], $this->arDepartments)) {
				$this->log['USERS_WITHOUT_DEP']++;
				echo '<b>Департамент не существует</b> ('.$u['ИДПодразделение'].')<br>';
				continue;
			}

			switch ($u['Пол']) {
				case 'Женский':
					$sex = 'F';
					break;
				case 'Мужской':
					$sex = 'M';
					break;
				default:
					$sex = '';
			}

			$userLogin = '';
			if (!empty($u['Windows_login'])) {
				//\\domain\i.ivanov
				$userLogin = str_replace('\\\\domain\\', '', trim($u['Windows_login']));
			}

			$arFields = array(
				'XML_ID'			=> $u['ИД'],
				'NAME'				=> $u['Имя'],
				'LAST_NAME'			=> $u['Фамилия'],
				'SECOND_NAME'		=> $u['Отчество'],
				'PERSONAL_GENDER'	=> $sex,
				'PERSONAL_BIRTHDAY'	=> (!empty($u['ДатаРождения']))? $u['ДатаРождения']:'',
				'WORK_COMPANY'		=> $u['Организация'],
				'WORK_POSITION'		=> $u['Должность'],
				'WORK_COUNTRY'		=> ($u['Страна'] == 'Россия')? '1':'',
				'WORK_STATE'		=> $u['Область'],
				'WORK_CITY'			=> $u['Город'],
				'WORK_ZIP'			=> $u['ПочтовыйИндекс'],
				'WORK_STREET'		=> (!empty($u['Улица']) && !empty($u['НомерДома']))? $u['Улица'].', '.$u['НомерДома']:'',
				'UF_DEPARTMENT'		=> array($this->arDepartments[$u['ИДПодразделение']]),
			);

			$obUser = new CUser();

			if (array_key_exists(mb_strtoupper($userLogin), $this->arUsersByLogin)) { //есть уже пользователь с таким логином
				//апдейт юзера
				if(!$obUser->Update($this->arUsersByLogin[$userLogin], $arFields)) {
					echo $obUser->LAST_ERROR.'<br><hr>';
					$this->errorCnt++;
					$this->log['UPDATING_ERRORS'][] = array('ERROR_TEXT'=>$obUser->LAST_ERROR,'FIELDS'=>$arFields);
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
					$this->messageBody .= 'Ошибка обновления пользователя по логину\r\n';
					$this->messageBody .= print_r($arFields,1);
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
				} else {
					$this->log['USERS_UPDATED']++;
					echo 'Успешно обновлён по логину<br>';
				}
			} elseif (array_key_exists($u['ИД'], $this->arUsersByXmlId)) {
				//апдейт
				if(!$obUser->Update($this->arUsersByXmlId[$u['ИД']], $arFields)) {
					echo $obUser->LAST_ERROR.'<br>';
					$this->log['UPDATING_ERRORS'][] = array('ERROR_TEXT'=>$obUser->LAST_ERROR,'FIELDS'=>$arFields);
					$this->errorCnt++;
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
					$this->messageBody .= 'Ошибка обновления пользователя по XML_ID'."\r\n";
					$this->messageBody .= print_r($arFields,1);
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
				} else {
					$this->log['USERS_UPDATED']++;
					echo 'Успешно обновлён по XML_ID<br>'/*'<pre>'.print_r($arFields,1).'</pre><hr>'*/;
				}
			} else {
				$fullName1 = mb_strtoupper($u['Имя'].'_'.$u['Фамилия']);
				$fullName2 = mb_strtoupper($u['Имя'].'_'.$u['Отчество'].'_'.$u['Фамилия']);
				$arSimilarUsers = array();
				foreach ($this->arUsersByName as $k => $id) {
					list($fullName, $uID) = explode('|',$k);
					if ($fullName == $fullName1 || $fullName == $fullName2) {
						$arSimilarUsers[] = $id;
					}
				}

				if (count($arSimilarUsers) === 1) {
					//update ADшного пользователя данными из 1С
					if(!$obUser->Update($arSimilarUsers[0], $arFields)) {
						echo $obUser->LAST_ERROR.'<br><hr>';
						$this->log['UPDATING_ERRORS'][] = array('ERROR_TEXT'=>$obUser->LAST_ERROR,'FIELDS'=>$arFields);
						$this->errorCnt++;
						$this->messageBody .= "\r\n".'================================================================'."\r\n";
						$this->messageBody .= "\r\n".'Ошибка обновления пользователя по ФИО'."\r\n";
						$this->messageBody .= print_r($arFields,1);
						$this->messageBody .= "\r\n".'================================================================'."\r\n";
					} else {
						$this->log['USERS_UPDATED']++;
						echo 'Успешно обновлен по ФИО<br>';
					}
					#continue;
				} elseif (count($arSimilarUsers) > 1) {
					$this->errorCnt++;
					$this->log['COLLIZIONS'][] = array(
						'SIMILAR' => $arSimilarUsers,
						'1CUSER'  => $u
					);

					$this->messageBody .= "\r\n".'================================================================'."\r\n";
					$this->messageBody .= "\r\n".'Невозможно однозначно сопоставить пользователей из 1С и на Портале'."\r\n";
					$this->messageBody .= "\r\n".'Пользователи с ID: '.implode(', ', $arSimilarUsers).' и пользователь из 1С: '.$fullName2.' ('.$u['ИД'].')'."\r\n";
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
				} else {
					//добавляем пользователя
					$arFields['LOGIN'] = ($userLogin != '')? $userLogin:$u['ИД'];
					$arFields['EMAIL'] = (strlen($userLogin) > 0)? $userLogin.'@mail.ru':'nomail@mail.ru';
					$pass = (strlen($userLogin) > 0)? randString(6):$this->defaultPass;
					$arFields['PASSWORD'] = $pass;
					$arFields['CONFIRM_PASSWORD'] = $pass;
					$arFields['GROUP_ID'] = array(2,3,4,11);

					$ID = $obUser->Add($arFields);
					if (intval($ID) > 0) {
						echo "Пользователь успешно добавлен #$ID.<br>";
						$this->log['USERS_ADDED']++;
						if ($userLogin == '') {
							$obUser->Update($ID, array('LOGIN'=>'user_'.$ID));
						}
						//формируем нормальный логин если логин не задан
						if ($arFields['EMAIL'] != 'nomail@mail.ru' && !empty($userLogin)) {
							CUser::SendUserInfo($ID, SITE_ID, 'Ваш профиль активирован!');
						}
					} else {
						$this->errorCnt++;
						$this->log['ADDING_ERRORS'][] = array('ERROR_TEXT'=>$obUser->LAST_ERROR,'FIELDS'=>$arFields);
						echo $obUser->LAST_ERROR.'<br>';
					}
				}
			}
			unset($ID);
			unset($obUser);
		}
		if ($this->errorCnt == 0) {
			$this->messageBody = 'Ошибок во время импорта не возникло';
		}
		$this->sendAdminMessage('Отчёт о импортe пользователей за '.date('d.m.Y H:i:s'), $this->messageBody);
		return $this->putLog('users');
	}

}

/**
 * Class CUsersStateHistoryImport1C
 * Import history states from 1C
 * @author Sergey Greznov <segey.greznov@gmail.com>
 */
class CUsersHistoryStatesImport1C extends CImport1C {

	/**
	 * Array with departments
	 * @var array Key - XML_ID, value - section ID
	 */
	private $arDepartments = array();

	/**
	 * Select departments from Intranet
	 * @return void
	 */
	private function selectDepartments () {
		$arFilter = array('IBLOCK_ID'=>self::depBlockID, 'GLOBAL_ACTIVE'=>'Y');
		$rsSect = CIBlockSection::GetList(array($by=>$order), $arFilter, false);

		while($arSect = $rsSect->Fetch()) {
			if (strlen(trim($arSect['XML_ID'])) > 0) {
				$this->arDepartments[trim($arSect['XML_ID'])] = $arSect['ID'];
			}
		}
	}

	public function doImport() {
		try {
			if (!CModule::IncludeModule('iblock') || !CModule::IncludeModule('main')) {
				throw new Exception('Bitrix modules "iblock" or "main" not included');
			}
		} catch (Exception $e) {
			echo '<h4>'.$e->getMessage().'</h4>';
		}

		$this->selectDepartments();
		$this->selectUsers();

		$this->log = array (
			'PROC_ROWS'				=> 0,
			'DEPARTMENT_NOT_EXISTS'	=> array(),
			'USERS_NOT_EXISTS'		=> array(),
			'ELEMENT_ADDED'			=> 0,
			'ELEMENT_ADDING_ERRORS'	=> array(),
		);

		if ($this->countXMLObjects == 0) {
			$this->messageBody .= 'Пустой (или не валидный) файл импорта.';
		}

		for ($i = 0; $i < $this->countXMLObjects; $i++) {
			$h = $this->arXMLObjects[$i];
			if (!array_key_exists($h['ИДПодразделение'], $this->arDepartments)) {
				echo '<b>Департамент не существует</b> ('.$h['ИДПодразделение'].')<br>';
				$this->log['DEPARTMENT_NOT_EXISTS'][] = $h['Подразделение'].' ('.$h['ИДПодразделение'].')';
				continue;
			}

			if (!array_key_exists($h['ИД'], $this->arUsersByXmlId)) {
				$this->errorCnt++;
				echo '<b>Пользователь не найден</b><br>';
				$this->messageBody .= "\r\n".'================================================================'."\r\n";
				$this->messageBody .= "\r\n".'Пользователь не найден: '.implode(', ', $h)."\r\n";
				$this->messageBody .= "\r\n".'================================================================'."\r\n";
				$this->log['USERS_NOT_EXISTS'][] = implode (', ',$h);
				continue;
			}

			//***********************************************
			$obElement 	= new CIBlockElement;
			$obUser 	= new CUser();

			$PROP = array(
				'USER'			=> $this->arUsersByXmlId[$h['ИД']],
				'USER_ACTIVE'	=> 'Y',
				'DEPARTMENT'	=> $this->arDepartments[$h['ИДПодразделение']],
				'POST'			=> $h['Должность'],
			);

			$rsUser = $obUser->GetByID($this->arUsersByXmlId[$h['ИД']]);
			$arUser = $rsUser->Fetch();

			$arFields = array(
				'ACTIVE_FROM'		=> date('d.m.Y H:i:s', strtotime($h['ДатаИзмененияСостояния'])),
				'NAME'				=> ' - '.$arUser['LAST_NAME'].' '.$arUser['NAME'],
				'IBLOCK_SECTION_ID'	=> false,
				'IBLOCK_ID'			=> 6,
				'ACTIVE'			=> 'Y'
			);

			switch ($h['Состояние']) {
				case 'Увольнение':
					$PROP['STATE'] = '32';
					$arFields['NAME'] = 'Уволен'.$arFields['NAME'];
					$arFields['PREVIEW_TEXT'] = 'Уволен';
					break;
				case 'Прием на работу':
					$PROP['STATE'] = '30';
					$arFields['NAME'] = 'Принят'.$arFields['NAME'];
					$arFields['PREVIEW_TEXT'] = 'Принят';
					break;
				case '':        //TODO: нет текста состояния "ПЕРЕВЕДЕН"
					$PROP['STATE'] = '31';
					$arFields['NAME'] = 'Переведен'.$arFields['NAME'];
					$arFields['PREVIEW_TEXT'] = 'Переведен';
					break;
			}

			if($histID = $obElement->Add($arFields)) {
				echo 'Элемент истории состояний добавлен'.'<br>';
				$this->log['ELEMENT_ADDED']++;
				CIBlockElement::SetPropertyValues($histID, 6, $PROP, false);
				if ($PROP['STATE'] == '32') {
					$fields = array(
						'ACTIVE' => 'N'
					);
					$obUser->Update($this->arUsersByXmlId[$h['ИД']], $fields);
				}
			} else {
				$this->errorCnt++;
				$this->messageBody .= "\r\n".'================================================================'."\r\n";
				$this->messageBody .= "\r\n".'Ошибка добавления элемента истории состояний: '.$obElement->LAST_ERROR."\r\n";
				$this->messageBody .= "\r\n".'================================================================'."\r\n";
				$this->log['ELEMENT_ADDING_ERRORS'][] = 'Ошибка добавления элемента истории состояний: '.$obElement->LAST_ERROR.' ('.implode(',',$arFields).')';
				echo "Ошибка добавления элемента истории состояний: ".$obElement->LAST_ERROR.'<br>';
			}

		}

		if ($this->errorCnt == 0) {
			$this->messageBody = 'Ошибок во время импорта не возникло';
		}
		$this->sendAdminMessage('Отчёт о импортe истории состояний за '.date('d.m.Y H:i:s'), $this->messageBody);
		return $this->putLog('users_hist');
	}

}

/**
 * Class CUsersStateHistoryImport1C
 * Import absence calendar from 1C
 * @author Sergey Greznov <segey.greznov@gmail.com>
 */
class CAbseneceCalendarImport1C extends CImport1C {

	/**
	 * Array of absence elements
	 * @var array Absence elements
	 */
	protected $arAbsence = array();

	/**
	 * IBlock ID with absences
	 * @const int
	 */
	const absBlockID = 3;

	/**
	 * Select absences from abs.calendar
	 * @return void
	 */
	protected function selectAbsences () {
		$arFilter = array('IBLOCK_ID'=>3, 'GLOBAL_ACTIVE'=>'Y');
		$rsElem = CIBlockElement::GetList(array(), $arFilter, false, false, array('ID','XML_ID'));
		while($arElem = $rsElem->Fetch()) {
			if (strlen(trim($arElem['XML_ID'])) == 36) {
				$this->arAbsence[trim($arElem['XML_ID'])] = $arElem['ID'];
			}
		}
	}

	public function doImport() {
		try {
			if (!CModule::IncludeModule('iblock') || !CModule::IncludeModule('main')) {
				throw new Exception('Bitrix modules "iblock" or "main" not included');
			}
		} catch (Exception $e) {
			echo '<h4>'.$e->getMessage().'</h4>';
		}

		$this->selectUsers();
		$this->selectAbsences();

		$this->log = array (
			'PROC_ROWS'					=> 0,
			'USERS_NOT_EXISTS'			=> array(),
			'ELEMENT_ADDED'				=> 0,
			'ELEMENT_ADDING_ERRORS'		=> array(),
			'ELEMENT_UPDATED'			=> 0,
			'ELEMENT_UPDATING_ERRORS'	=> array(),
		);

		if ($this->countXMLObjects == 0) {
			$this->messageBody .= 'Пустой (или не валидный) файл импорта.';
		}

		for ($i = 0; $i < $this->countXMLObjects; $i++) {
			$h = $this->arXMLObjects[$i];

			if (!array_key_exists($h['ИД'], $this->arUsersByXmlId)) {
				$this->errorCnt++;
				echo '<b>Пользователь не найден</b><br>';
				$this->messageBody .= "\r\n".'================================================================'."\r\n";
				$this->messageBody .= "\r\n".'Пользователь не найден: '.implode(', ', $h)."\r\n";
				$this->messageBody .= "\r\n".'================================================================'."\r\n";
				$this->log['USERS_NOT_EXISTS'][] = implode (', ',$h);
				continue;
			}

			//***********************************************
			$obElement 	= new CIBlockElement;

			$absType = 14;#OTHER

			switch (trim($h['ТипОтсутствия'])) {
				case 'Болеет':#LEAVESICK
				case 'Болеет из-за травмы на производстве':#LEAVESICK
					$absType = 10;
					break;
				case 'В командировке':#ASSIGNMENT
					$absType = 9;
					break;
				case 'В ежегодном отпуске':#VACATION
					$absType = 8;
					break;
				case 'В отпуске по беременности и родам':#LEAVEMATERINITY
					$absType = 11;
					break;
				case 'В отпуске по уходу за ребенком':#LEAVEPREGNAT
					$absType = 310;
					break;
				case 'Прогулы':#UNKNOWN
					$absType = 13;
					break;
				case 'В отпуске без сохранения зарплаты':#LEAVEUNPAYED
					$absType = 12;
					break;
				default:#OTHER
					$absType = 14;
					break;
			}

			$PROP = array(
				'USER'			=> $this->arUsersByXmlId[$h['ИД']],
				'FINISH_STATE'	=> '',
				'STATE'			=> '',
				'ABSENCE_TYPE'	=> $absType,
			);

			$arFields = array(
				'ACTIVE_FROM'		=> date('d.m.Y H:i:s', strtotime($h['ДатаС'])),
				'ACTIVE_TO'			=> date('d.m.Y H:i:s', strtotime($h['ДатаПо'])),
				'NAME'				=> $h['ТипОтсутствия'],
				'IBLOCK_SECTION_ID'	=> false,
				'IBLOCK_ID'			=> 3,
				'ACTIVE'			=> 'Y',
				'XML_ID'			=> $h['ИдДокумента']
			);

			if (array_key_exists($h['ИдДокумента'], $this->arAbsence)) {
				//update
				if($obElement->Update($this->arAbsence[$h['ИдДокумента']], $arFields)) {
					echo 'Элемент календаря отсутствий обновлен'.'<br>';
					$this->log['ELEMENT_UPDATED']++;
					CIBlockElement::SetPropertyValues($this->arAbsence[$h['ИдДокумента']], self::absBlockID, $PROP, false);
				} else {
					$this->errorCnt++;
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
					$this->messageBody .= "\r\n".'Ошибка обновления элемента календаря отсутствий: '.$obElement->LAST_ERROR."\r\n";
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
					$this->log['ELEMENT_UPDATING_ERRORS'][] = 'Ошибка обновления элемента календаря отсутствий: '.$obElement->LAST_ERROR.' ('.implode(',',$arFields).')';
					echo "Ошибка обновления элемента календаря отсутствий: ".$obElement->LAST_ERROR.'<br>';
				}
			} else {
				//insert
				if($elemID = $obElement->Add($arFields)) {
					echo 'Элемент истории состояний добавлен'.'<br>';
					$this->log['ELEMENT_ADDED']++;
					CIBlockElement::SetPropertyValues($elemID, self::absBlockID, $PROP, false);
				} else {
					$this->errorCnt++;
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
					$this->messageBody .= "\r\n".'Ошибка добавления элемента истории состояний: '.$obElement->LAST_ERROR."\r\n";
					$this->messageBody .= "\r\n".'================================================================'."\r\n";
					$this->log['ELEMENT_ADDING_ERRORS'][] = 'Ошибка добавления элемента истории состояний: '.$obElement->LAST_ERROR.' ('.implode(',',$arFields).')';
					echo "Ошибка добавления элемента истории состояний: ".$obElement->LAST_ERROR.'<br>';
				}
			}
		}

		if ($this->errorCnt == 0) {
			$this->messageBody = 'Ошибок во время импорта не возникло';
		}
		$this->sendAdminMessage('Отчёт о импортe календаря отсутствий за '.date('d.m.Y H:i:s'), $this->messageBody);
		return $this->putLog('absence');
	}

}


?>