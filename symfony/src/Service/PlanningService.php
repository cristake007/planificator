<?php

namespace App\Service;

use DateInterval;
use DateTimeImmutable;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use RuntimeException;
use SimpleXMLElement;
use SplFileObject;
use ZipArchive;

class PlanningService
{
    private const MONTH_NAMES = [1 => 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

    public function generateSchedule(string $path, string $filename, int $year, int $randomness, array $months, array $holidays): array
    {
        $rows = $this->readCourseRows($path, $filename);
        $holidaySet = array_flip(array_filter(array_map(fn ($h) => $this->parseHoliday($h), $holidays)));
        $schedule = [];
        $unscheduled = [];

        foreach ($months as $monthRaw) {
            $month = (int) $monthRaw;
            $scheduledStarts = [];
            $monthUnscheduled = [];

            foreach ($rows as $row) {
                $duration = (int) $row['duration'];
                $available = $this->getAvailableStartDays($year, $month, $duration, $holidaySet);
                if (!$available) {
                    $monthUnscheduled[] = $row['Title'];
                    continue;
                }

                $minGap = max(1, 11 - $randomness);
                $filtered = array_values(array_filter($available, function (DateTimeImmutable $date) use ($scheduledStarts, $minGap) {
                    foreach ($scheduledStarts as $scheduled) {
                        if (abs($date->diff($scheduled)->days) < $minGap) {
                            return false;
                        }
                    }
                    return true;
                }));

                $candidates = $filtered ?: $available;
                $startDate = $this->pickStartDate($candidates, $duration, $randomness);
                if (!$startDate) {
                    $monthUnscheduled[] = $row['Title'];
                    continue;
                }

                $schedule[] = [
                    'Title' => $row['Title'],
                    'Permalink' => $row['Permalink'],
                    'Durata Curs' => $row['Durata Curs'],
                    'investitie' => $row['investitie'],
                    'date_range' => $this->formatDateRange($startDate, $duration, $holidaySet),
                    'month' => $month,
                    'original_order' => $row['original_order'],
                ];
                $scheduledStarts[] = $startDate;
            }

            if ($monthUnscheduled) {
                sort($monthUnscheduled);
                $unscheduled[$month] = $monthUnscheduled;
            }
        }

        if ($unscheduled) {
            $messages = [];
            foreach ($unscheduled as $month => $courses) {
                $messages[] = self::MONTH_NAMES[$month] . ': ' . implode(', ', $courses);
            }
            return [
                'success' => false,
                'error' => "Unable to schedule all courses with the current constraints. The following courses had no available dates:\n" . implode("\n", $messages),
                'unscheduled_courses' => $unscheduled,
                'status' => 400,
            ];
        }

        usort($schedule, fn ($a, $b) => $a['original_order'] <=> $b['original_order']);
        return ['success' => true, 'schedule' => $schedule];
    }

    public function exportSchedule(array $schedule, int $year, array $holidays): array
    {
        $courses = [];
        foreach ($schedule as $item) {
            $title = (string) ($item['Title'] ?? '');
            if ($title === '') {
                continue;
            }
            if (!isset($courses[$title])) {
                $courses[$title] = [
                    'Title' => $title,
                    'Permalink' => (string) ($item['Permalink'] ?? ''),
                    'Durata Curs' => (string) ($item['Durata Curs'] ?? ''),
                    'investitie' => (string) ($item['investitie'] ?? ''),
                    'months' => array_fill(1, 12, ''),
                ];
            }
            $month = (int) ($item['month'] ?? 0);
            if ($month >= 1 && $month <= 12) {
                $courses[$title]['months'][$month] = (string) ($item['date_range'] ?? '');
            }
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Schedule');
        $headers = ['Title', 'Permalink', 'Durata Curs', 'investitie', ...array_values(self::MONTH_NAMES)];
        $sheet->fromArray($headers, null, 'A1');

        $rowIndex = 2;
        foreach ($courses as $course) {
            $line = [$course['Title'], $course['Permalink'], $course['Durata Curs'], $course['investitie']];
            for ($m = 1; $m <= 12; $m++) {
                $line[] = $course['months'][$m] ?? '';
            }
            $sheet->fromArray($line, null, 'A' . $rowIndex++);
        }

        if ($holidays) {
            $holidaySheet = $spreadsheet->createSheet();
            $holidaySheet->setTitle('Holidays');
            $holidaySheet->setCellValue('A1', 'Holiday Date');
            $i = 2;
            foreach ($holidays as $holiday) {
                $holidaySheet->setCellValue('A' . $i++, (string) $holiday);
            }
        }

        $tmp = tmpfile();
        $meta = stream_get_meta_data($tmp);
        $writer = new Xlsx($spreadsheet);
        $writer->save($meta['uri']);
        $bytes = file_get_contents($meta['uri']);
        fclose($tmp);

        return [
            'filename' => "course_schedule_{$year}.xlsx",
            'content_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'bytes' => $bytes,
        ];
    }

    public function formatXml(string $path, string $filename): array
    {
        $rows = $this->readTabularRows($path, $filename);
        $columns = $rows[0] ?? [];
        $normalized = [];
        foreach (array_keys($columns) as $name) {
            $normalized[$this->normalizeColumnName((string) $name)] = $name;
        }
        if (!isset($normalized['title'], $normalized['permalink'])) {
            throw new RuntimeException('Missing required columns: Title, Permalink');
        }

        $monthAliases = ['january','february','march','april','may','june','july','august','september','october','november','december',
            'luna 1','luna 2','luna 3','luna 4','luna 5','luna 6','luna 7','luna 8','luna 9','luna 10','luna 11','luna 12'];
        $monthColumns = [];
        foreach ($normalized as $lower => $original) {
            if (in_array($lower, $monthAliases, true)) {
                $monthColumns[] = $original;
            }
        }
        if (!$monthColumns) {
            throw new RuntimeException('No supported date columns found.');
        }

        $events = [];
        foreach ($rows as $row) {
            $title = trim((string) ($row[$normalized['title']] ?? ''));
            $permalink = trim((string) ($row[$normalized['permalink']] ?? ''));
            if ($title === '') {
                continue;
            }
            foreach ($monthColumns as $column) {
                $value = trim((string) ($row[$column] ?? ''));
                if ($value !== '' && strtolower($value) !== 'nan') {
                    $events[] = ['course_name' => $title, 'date_range' => $value, 'permalink' => $permalink];
                }
            }
        }

        $xml = $this->buildXml($events);
        return ['filename' => 'formatted_courses_' . date('Y') . '.xml', 'content_type' => 'application/xml', 'bytes' => $xml];
    }

    public function convertWord(string $wordPath, string $schedulePath, string $scheduleFilename): array
    {
        $rows = $this->readTabularRows($schedulePath, $scheduleFilename);
        if (!$rows) {
            throw new RuntimeException('No valid rows in schedule file.');
        }

        $headers = array_keys($rows[0]);
        $normalizedColumns = [];
        foreach ($headers as $header) {
            $normalizedColumns[$this->normalizeColumnName((string) $header)] = $header;
        }
        if (!isset($normalizedColumns['title'])) {
            throw new RuntimeException('Input file must contain a "Title" column');
        }

        $monthColumns = [];
        foreach (array_values(self::MONTH_NAMES) as $month) {
            $key = strtolower($month);
            if (isset($normalizedColumns[$key])) {
                $monthColumns[] = $normalizedColumns[$key];
            }
        }
        if (!$monthColumns) {
            throw new RuntimeException('Input file must contain month columns (January-December)');
        }

        $scheduleRows = [];
        foreach ($rows as $row) {
            $title = trim((string) ($row[$normalizedColumns['title']] ?? ''));
            if ($title === '') continue;
            $dates = [];
            foreach ($monthColumns as $col) {
                $value = trim((string) ($row[$col] ?? ''));
                if ($value !== '' && strtolower($value) !== 'nan') {
                    $dates[] = $value;
                }
                if (count($dates) === 3) break;
            }
            while (count($dates) < 3) $dates[] = '';
            $scheduleRows[] = ['title' => $title, 'normalized' => $this->normalize($title), 'dates' => $dates];
        }

        $zip = new ZipArchive();
        if ($zip->open($wordPath) !== true) {
            throw new RuntimeException('Failed to open Word file.');
        }

        $xml = $zip->getFromName('word/document.xml');
        if ($xml === false) {
            $zip->close();
            throw new RuntimeException('Invalid Word file content.');
        }

        $dom = new \DOMDocument();
        $dom->loadXML($xml);
        $xp = new \DOMXPath($dom);
        $xp->registerNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');

        foreach ($xp->query('//w:tbl/w:tr') as $tr) {
            $cells = $xp->query('./w:tc', $tr);
            if ($cells->length < 6) continue;
            $title = trim($this->extractCellText($cells->item(0), $xp));
            if ($title === '') continue;
            $match = $this->bestMatch($title, $scheduleRows);
            if ($match === null) continue;
            foreach ([3, 4, 5] as $idx => $targetCellIndex) {
                if ($cells->length > $targetCellIndex) {
                    $this->setCellText($cells->item($targetCellIndex), $match['dates'][$idx], $dom, $xp);
                }
            }
        }

        $zip->addFromString('word/document.xml', $dom->saveXML());
        $zip->close();

        return [
            'filename' => 'matched_courses.docx',
            'content_type' => 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'bytes' => file_get_contents($wordPath),
        ];
    }

    private function parseHoliday(string $value): ?string
    {
        $dt = DateTimeImmutable::createFromFormat('d.m.Y', trim($value));
        return $dt ? $dt->format('Y-m-d') : null;
    }

    private function readCourseRows(string $path, string $filename): array
    {
        $rows = $this->readTabularRows($path, $filename);
        if (!$rows) throw new RuntimeException('No data rows found.');

        $sample = $rows[0];
        $normMap = [];
        foreach (array_keys($sample) as $col) {
            $normMap[$this->normalizeColumnName((string) $col)] = $col;
        }
        foreach (['title' => 'Title', 'durata curs' => 'Durata Curs', 'permalink' => 'Permalink'] as $k => $v) {
            if (!isset($normMap[$k])) throw new RuntimeException("Missing required columns: {$v}");
        }

        $out = [];
        foreach ($rows as $i => $row) {
            $title = trim((string) ($row[$normMap['title']] ?? ''));
            $durata = trim((string) ($row[$normMap['durata curs']] ?? ''));
            if ($title === '' || $durata === '') continue;
            if (!preg_match('/(\d+)/', $durata, $m)) {
                throw new RuntimeException('Invalid values in "Durata Curs".');
            }
            $out[] = [
                'Title' => $title,
                'Permalink' => trim((string) ($row[$normMap['permalink']] ?? '')),
                'Durata Curs' => $durata,
                'duration' => (int) $m[1],
                'investitie' => isset($normMap['investitie']) ? trim((string) ($row[$normMap['investitie']] ?? '')) : '',
                'original_order' => $i,
            ];
        }

        return $out;
    }

    private function readTabularRows(string $path, string $filename): array
    {
        $extension = strtolower(pathinfo($filename, PATHINFO_EXTENSION));
        if ($extension === 'csv') {
            $lines = file($path, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
            if (!$lines) {
                return [];
            }

            $delimiter = $this->detectCsvDelimiter($lines[0]);
            $file = new SplFileObject($path, 'r');
            $headers = $file->fgetcsv($delimiter, '"', '\\');
            if (!is_array($headers)) {
                return [];
            }

            $headers = array_map(fn ($value) => trim((string) $value), $headers);
            $headerCount = count($headers);
            if ($headerCount === 0) {
                return [];
            }

            $rows = [];
            while (!$file->eof()) {
                $values = $file->fgetcsv($delimiter, '"', '\\');
                if (!is_array($values) || $values === [null]) {
                    continue;
                }

                $normalizedValues = array_slice(array_pad($values, $headerCount, ''), 0, $headerCount);
                $rows[] = array_combine($headers, $normalizedValues);
            }

            return $rows;
        }

        $spreadsheet = IOFactory::load($path);
        $sheet = $spreadsheet->getActiveSheet();
        $raw = $sheet->toArray('', true, true, false);
        if (!$raw) return [];
        $headers = array_map(fn ($v) => trim((string) $v), array_shift($raw));
        $rows = [];
        foreach ($raw as $line) {
            if (implode('', array_map('strval', $line)) === '') continue;
            $assoc = [];
            foreach ($headers as $idx => $header) {
                if ($header === '') continue;
                $assoc[$header] = isset($line[$idx]) ? trim((string) $line[$idx]) : '';
            }
            if ($assoc) $rows[] = $assoc;
        }
        return $rows;
    }

    private function detectCsvDelimiter(string $firstLine): string
    {
        $candidates = ['@', ';', ',', "\t", '|'];
        $bestDelimiter = ',';
        $bestCount = -1;

        foreach ($candidates as $candidate) {
            $count = substr_count($firstLine, $candidate);
            if ($count > $bestCount) {
                $bestCount = $count;
                $bestDelimiter = $candidate;
            }
        }

        return $bestDelimiter;
    }

    private function normalizeColumnName(string $name): string
    {
        $name = trim($name);
        if (str_starts_with($name, "\xEF\xBB\xBF")) {
            $name = substr($name, 3);
        }

        return strtolower(trim($name));
    }

    private function getAvailableStartDays(int $year, int $month, int $duration, array $holidaySet): array
    {
        $start = new DateTimeImmutable("{$year}-{$month}-01");
        $end = $start->modify('last day of this month');
        $dates = [];
        for ($d = $start; $d <= $end; $d = $d->add(new DateInterval('P1D'))) {
            if ($this->canSchedule($d, $duration, $holidaySet)) {
                $dates[] = $d;
            }
        }
        return $dates;
    }

    private function canSchedule(DateTimeImmutable $startDate, int $duration, array $holidaySet): bool
    {
        if (!$this->isBusinessDay($startDate, $holidaySet)) return false;
        $current = $startDate;
        $businessDays = 0;
        $allowCrossPeriod = $duration > 5;
        $weekStart = $startDate->modify('monday this week');

        while ($businessDays < $duration) {
            if (!$allowCrossPeriod) {
                if ($current >= $weekStart->add(new DateInterval('P5D'))) return false;
                if ($current->format('m') !== $startDate->format('m')) return false;
            }
            if ($this->isBusinessDay($current, $holidaySet)) $businessDays++;
            $current = $current->add(new DateInterval('P1D'));
        }
        return true;
    }

    private function isBusinessDay(DateTimeImmutable $date, array $holidaySet): bool
    {
        $dayOfWeek = (int) $date->format('N');
        if ($dayOfWeek >= 6) return false;
        return !isset($holidaySet[$date->format('Y-m-d')]);
    }

    private function pickStartDate(array $dates, int $duration, int $randomness): ?DateTimeImmutable
    {
        if (!$dates) return null;
        if ($duration > 5) {
            usort($dates, fn ($a, $b) => $a <=> $b);
            return $dates[0];
        }
        if ($randomness > 7) {
            return $dates[array_rand($dates)];
        }
        return $dates[0];
    }

    private function formatDateRange(DateTimeImmutable $startDate, int $duration, array $holidaySet): string
    {
        if ($duration === 1) return $startDate->format('d.m.Y');
        $businessDays = 0;
        $current = $startDate;
        while ($businessDays < $duration) {
            if ($this->isBusinessDay($current, $holidaySet)) $businessDays++;
            if ($businessDays < $duration) $current = $current->add(new DateInterval('P1D'));
        }
        return $startDate->format('d') . '-' . $current->format('d.m.Y');
    }

    private function buildXml(array $events): string
    {
        $root = new SimpleXMLElement('<events/>');
        $grouped = [];
        foreach ($events as $event) {
            $grouped[$event['course_name']][] = $event;
        }
        $eventId = 20000;
        foreach ($grouped as $courseName => $courseEvents) {
            foreach ($courseEvents as $periodIdx => $event) {
                [$startDate, $endDate] = $this->parseDateRange($event['date_range']);
                $eventId++;
                $item = $root->addChild('item');
                $item->addChild('ID', (string) $eventId);
                $item->addChild('title', htmlspecialchars($courseName));
                $item->addChild('content', '');
                $meta = $item->addChild('meta');
                $meta->addChild('mec_more_info_title', 'perioada ' . ($periodIdx + 1));
                $meta->addChild('mec_read_more', htmlspecialchars($event['permalink']));
                $meta->addChild('mec_allday', '1');
                $meta->addChild('mec_start_date', $startDate);
                $meta->addChild('mec_end_date', $endDate);
            }
        }
        return $root->asXML() ?: '<events/>';
    }

    private function parseDateRange(string $dateRange): array
    {
        $normalized = trim($dateRange);
        if (preg_match('/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/', $normalized, $m)) {
            $iso = sprintf('%s-%02d-%02d', $m[3], $m[2], $m[1]);
            return [$iso, $iso];
        }
        if (preg_match('/^(\d{1,2})\s*-\s*(\d{1,2})\.(\d{1,2})\.(\d{4})$/', $normalized, $m)) {
            return [sprintf('%s-%02d-%02d', $m[4], $m[3], $m[1]), sprintf('%s-%02d-%02d', $m[4], $m[3], $m[2])];
        }
        if (preg_match('/^(\d{1,2})\.(\d{1,2})\s*-\s*(\d{1,2})\.(\d{1,2})\.(\d{4})$/', $normalized, $m)) {
            return [sprintf('%s-%02d-%02d', $m[5], $m[2], $m[1]), sprintf('%s-%02d-%02d', $m[5], $m[4], $m[3])];
        }
        throw new RuntimeException("Unsupported date format: {$dateRange}");
    }

    private function normalize(string $value): string
    {
        $value = strtolower(trim($value));
        $value = preg_replace('/[^\w\s]/u', ' ', $value);
        return trim(preg_replace('/\s+/', ' ', (string) $value));
    }

    private function similarity(string $a, string $b): float
    {
        similar_text($a, $b, $percent);
        $at = array_filter(explode(' ', $a));
        $bt = array_filter(explode(' ', $b));
        $overlap = $bt ? (count(array_intersect($at, $bt)) / count($bt) * 100) : 0;
        return 0.7 * $percent + 0.3 * $overlap;
    }

    private function bestMatch(string $wordTitle, array $scheduleRows): ?array
    {
        $target = $this->normalize($wordTitle);
        if ($target === '') return null;
        $best = null;
        $bestScore = 0;
        foreach ($scheduleRows as $row) {
            $score = $this->similarity($target, $row['normalized']);
            if ($score > $bestScore) {
                $bestScore = $score;
                $best = $row;
            }
        }
        return $bestScore >= 70 ? $best : null;
    }

    private function extractCellText(\DOMNode $cell, \DOMXPath $xp): string
    {
        $parts = [];
        foreach ($xp->query('.//w:t', $cell) as $node) {
            $parts[] = $node->textContent;
        }
        return trim(implode('', $parts));
    }

    private function setCellText(\DOMNode $cell, string $value, \DOMDocument $dom, \DOMXPath $xp): void
    {
        foreach ($xp->query('.//w:p', $cell) as $paragraph) {
            $paragraph->parentNode?->removeChild($paragraph);
        }
        $p = $dom->createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:p');
        $r = $dom->createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
        $t = $dom->createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:t');
        $t->appendChild($dom->createTextNode($value));
        $r->appendChild($t);
        $p->appendChild($r);
        $cell->appendChild($p);
    }
}
