<?php
ini_set('display_errors', 1);
error_reporting(E_ALL);
date_default_timezone_set("Europe/Moscow");

// === Определяем дату и границы недели ===
$todayDate = new DateTime();   // всегда сегодня
$monday = clone $todayDate;
$monday->modify('monday this week');
$sunday = clone $monday;
$sunday->modify('sunday this week');

$xml = simplexml_load_file(__DIR__ . "/rasp3.xml");

// === справочники ===
$groups = [];
foreach ($xml->group->grp as $g) $groups[(string)$g['id']] = (string)$g['name'];
$rooms = [];
foreach ($xml->room->rom as $r) $rooms[(string)$r['id']] = (string)$r['nazv'];
$preds = [];
foreach ($xml->predm->pred as $p) $preds[(string)$p['id']] = (string)$p['full'];
$prepods = [];
foreach ($xml->prepod->prep as $p) {
    $fio = $p['fam'];
    if ($p['name'] !== "") $fio .= " " . mb_substr($p['name'],0,1,"UTF-8") . ".";
    if ($p['otch'] !== "") $fio .= mb_substr($p['otch'],0,1,"UTF-8") . ".";
    $prepods[(string)$p['id']] = $fio;
}

// === индексы ===
$ldById = [];
foreach ($xml->xpath('//ld') as $ld) $ldById[(string)$ld['loadsid']] = $ld;
$gpByInn = [];
foreach ($xml->xpath('//gp') as $gp) {
    $inn = (string)$gp['grpinn'];
    $gpByInn[$inn][] = $gp;
}

function splitIds($s) {
    $parts = preg_split('/[,\s;]+/u', (string)$s, -1, PREG_SPLIT_NO_EMPTY);
    return array_map('trim', $parts);
}

// === функция для цвета по предмету ===
function subjectColor($subject) {
    $hash = crc32($subject);
    $hue = $hash % 360; // оттенок
    return "hsl($hue, 70%, 85%)"; // пастельный цвет
}

// === собираем расписание ===
$table = [];

foreach ($xml->xpath('//sh') as $sh) {
    $beginDate = DateTime::createFromFormat('d.m.Y', (string)$sh['shedbegin_date']);
    $endDate   = DateTime::createFromFormat('d.m.Y', (string)$sh['shedend_date']);
    if (!$beginDate || !$endDate) continue;

    // фильтр: пересечение с текущей неделей
    if ($endDate < $monday || $beginDate > $sunday) continue;

    $ld = $ldById[(string)$sh['shedload_id']] ?? null;
    if (!$ld) continue;

    $klassList = splitIds($ld['loadsklass_id_list'] ?? "");
    if (!$klassList) continue;

    $day   = (int)$sh['shedday'];
    $hour  = (int)$sh['shedhour'];
    $pair  = (int)ceil($hour / 2);
    $room  = $rooms[(string)$sh['shedroom_id']] ?? "?";

    $loadsnum = trim((string)($ld['loadsnum'] ?? ""));
    $loadsgrpn = (int)($ld['loadsgrpn'] ?? 0);
    $shedgroup = (int)($sh['shedgroup'] ?? 0);

    $gps = $gpByInn[$loadsnum] ?? [];
    foreach ($klassList as $grpId) {
        $gname = $groups[$grpId] ?? ("Гр.".$grpId);
        foreach ($gps as $gp) {
            $teacher = $prepods[(string)$gp['grpteacher_id']] ?? "[?]";
            $subject = $preds[(string)$gp['grpsubject_id']] ?? "[?]";

            // подгруппы
            $subgroup = 0;
            if ($loadsgrpn >= 3) {
                if ($shedgroup === 0) $subgroup = 1;
                elseif ($shedgroup === 1) $subgroup = 2;
                elseif ($shedgroup === 2) $subgroup = 3;
            }

            // ключ для пары и подгруппы
            $key = $day.'-'.$pair.'-'.$gname.'-'.$subgroup;
            if (!isset($table[$key])) {
                $table[$key] = [];
            }

            // объединение по предмету
            $found = false;
            foreach ($table[$key] as &$rec) {
                if ($rec['subject'] === $subject) {
                    if (!in_array($teacher, $rec['teachers'])) {
                        $rec['teachers'][] = $teacher;
                    }
                    $rec['rooms'][] = $room;
                    $found = true;
                    break;
                }
            }
            unset($rec);

            if (!$found) {
                $table[$key][] = [
                    'subject'  => $subject,
                    'teachers' => [$teacher],
                    'rooms'    => [$room],
                    'dates'    => $beginDate->format('d.m.Y') . " - " . $endDate->format('d.m.Y')
                ];
            }
        }
    }
}

// === вывод ===
$days = [1=>"Пн",2=>"Вт",3=>"Ср",4=>"Чт",5=>"Пт",6=>"Сб"];
$groupNames = array_values($groups);

// макс. число пар
$maxPair = 0;
foreach ($table as $key=>$recs) {
    [$d,$p] = explode('-',$key);
    if ($p > $maxPair) $maxPair = $p;
}
if ($maxPair == 0) $maxPair = 6;
?>
<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8">
<title>Расписание всех групп (неделя <?= $monday->format('d.m.Y') ?> - <?= $sunday->format('d.m.Y') ?>)</title>
<style>
body { background:#fff; color:#000; font-family:Arial,sans-serif; margin:0; }
table { border-collapse:collapse; width:100%; font-size:13px; }
th, td { border:1px solid #555; padding:4px; text-align:left; vertical-align:top; }
th { background:#eee; text-align:center; }
.sub { margin-top:4px; padding-top:4px; font-size:12px; }
small { color:#333; }
.subject-box { margin:2px 0; padding:2px; border-radius:4px; }
</style>
</head>
<body>
<h2 style="text-align:center;">Расписание всех групп (неделя <?= $monday->format('d.m.Y') ?> - <?= $sunday->format('d.m.Y') ?>)</h2>
<table>
<tr>
  <th>День</th><th>Пара</th>
  <?php foreach ($groupNames as $gname): ?>
    <th><?= htmlspecialchars($gname) ?></th>
  <?php endforeach; ?>
</tr>
<?php foreach ($days as $dnum=>$dname): ?>
  <?php for ($p=1; $p<=$maxPair; $p++): ?>
    <tr>
      <td><?= $dname ?></td>
      <td><?= $p ?></td>
      <?php foreach ($groupNames as $gname): ?>
        <td>
          <?php
          foreach ([0,1,2,3] as $sg) {
              $key = $dnum.'-'.$p.'-'.$gname.'-'.$sg;
              $cell = $table[$key] ?? [];
              if (!$cell) continue;

              if ($sg > 0) echo "<div><b>{$sg} подгр.:</b><br>";

              foreach ($cell as $rec) {
                  $color = subjectColor($rec['subject']);
                  echo "<div class='subject-box' style='background:".$color."'>";
                  echo htmlspecialchars($rec['subject'])." — ".htmlspecialchars(implode(', ', $rec['teachers'])).
                       " (".htmlspecialchars(implode(', ', array_unique($rec['rooms']))).")<br>".
                       "<small>".htmlspecialchars($rec['dates'])."</small>";
                  echo "</div>";
              }

              if ($sg > 0) echo "</div>";
          }
          ?>
        </td>
      <?php endforeach; ?>
    </tr>
  <?php endfor; ?>
<?php endforeach; ?>
</table>
</body>
</html>
