<?php
/**
 * Plugin Name: Daily_Positive_Word (JP/EN CSV)
 * Description: uploads内のCSV（jp,en）から日替わりで日本語＋英語を表示。フロントページに自動表示＋ショートコード[daily_positive_word]
 * Version: 1.1.0
 */

if (!defined('ABSPATH')) exit;

class Daily_Positive_Word {
    const OPT_SOURCE = 'soraumi_quotes_csv_source';
    const OPT_CACHE  = 'soraumi_quotes_csv_cache';

    public static function init() {
        add_action('admin_menu', [__CLASS__, 'admin_menu']);
        add_action('admin_init', [__CLASS__, 'admin_init']);
        add_shortcode('daily_positive_word', [__CLASS__, 'shortcode']);
        add_filter('the_content', [__CLASS__, 'inject_front_page_quote'], 9);
    }

    public static function admin_menu() {
        add_options_page(
            'Daily Quotes CSV (JP/EN)',
            'Daily Quotes CSV',
            'manage_options',
            'soraumi-daily-quotes-csv',
            [__CLASS__, 'settings_page']
        );
    }

    public static function admin_init() {
        register_setting('soraumi_daily_quotes', self::OPT_SOURCE, [
            'type' => 'string',
            'sanitize_callback' => function($v){ return is_string($v) ? trim($v) : ''; },
            'default' => '',
        ]);
    }

    public static function settings_page() {
        if (!current_user_can('manage_options')) return;
        $source = get_option(self::OPT_SOURCE, '');
        ?>
        <div class="wrap">
            <h1>Daily Quotes CSV（JP/EN）設定</h1>
            <form method="post" action="options.php">
                <?php settings_fields('soraumi_daily_quotes'); ?>
                <table class="form-table" role="presentation">
                    <tr>
                        <th scope="row"><label for="<?php echo esc_attr(self::OPT_SOURCE); ?>">CSVの場所</label></th>
                        <td>
                            <input
                                type="text"
                                class="regular-text"
                                id="<?php echo esc_attr(self::OPT_SOURCE); ?>"
                                name="<?php echo esc_attr(self::OPT_SOURCE); ?>"
                                value="<?php echo esc_attr($source); ?>"
                                placeholder="例）添付ID（1234） または uploadsのURL"
                            />
                            <p class="description">
                                推奨：メディアライブラリの「添付ID」を入力（最も確実）。<br>
                                CSVは 1行=1件、列は <code>jp,en</code>（ヘッダーあり推奨）を想定します。
                            </p>
                        </td>
                    </tr>
                </table>
                <?php submit_button('保存'); ?>
            </form>

            <hr>
            <h2>表示</h2>
            <ul>
                <li>フロントページ：本文の先頭に自動で表示します</li>
                <li>ショートコード：<code>[daily_positive_word]</code></li>
            </ul>
        </div>
        <?php
    }

    private static function resolve_local_path($source) {
        $source = trim((string)$source);
        if ($source === '') return '';

        // 添付ID
        if (ctype_digit($source)) {
            $path = get_attached_file((int)$source);
            return (is_string($path) && $path !== '') ? $path : '';
        }

        // uploads URL -> 実ファイルパス
        if (filter_var($source, FILTER_VALIDATE_URL)) {
            $uploads = wp_upload_dir();
            if (strpos($source, $uploads['baseurl']) === 0) {
                $relative = substr($source, strlen($uploads['baseurl']));
                return $uploads['basedir'] . $relative;
            }
            return '';
        }

        // パス直接
        return $source;
    }

    private static function load_rows() {
        $source = get_option(self::OPT_SOURCE, '');
        $path = self::resolve_local_path($source);

        if (!$path || !file_exists($path) || !is_readable($path)) {
            return [[], 'CSVが見つからないか、読み取れません。設定（添付ID/URL）と権限を確認してください。'];
        }

        // 更新時刻でキャッシュ
        $mtime = @filemtime($path);
        $cache = get_option(self::OPT_CACHE, []);
        if (is_array($cache)
            && isset($cache['mtime'], $cache['rows'])
            && (int)$cache['mtime'] === (int)$mtime
            && is_array($cache['rows'])
        ) {
            return [$cache['rows'], ''];
        }

        $rows = [];
        $handle = fopen($path, 'rb');
        if (!$handle) return [[], 'CSVを開けませんでした。'];

        // UTF-8 BOM除去（ファイル先頭）
        $firstBytes = fread($handle, 3);
        if ($firstBytes !== "\xEF\xBB\xBF") rewind($handle);

        $line = 0;
        $headerMap = ['jp'=>0,'en'=>1];

        while (($cols = fgetcsv($handle, 0, ",")) !== false) {
            $line++;
            if (!is_array($cols) || count($cols) === 0) continue;

            $cols = array_map(function($v){
                $v = is_string($v) ? trim($v) : '';
                return $v;
            }, $cols);

            // 1行目がヘッダー（jp,en）ならスキップ＆位置確定
            if ($line === 1) {
                $lower = array_map(function($v){
                    return mb_strtolower(trim($v));
                }, $cols);

                if (in_array('jp', $lower, true) && in_array('en', $lower, true)) {
                    $headerMap = [
                        'jp' => array_search('jp', $lower, true),
                        'en' => array_search('en', $lower, true),
                    ];
                    continue;
                }
            }

            $jp = isset($cols[$headerMap['jp']]) ? trim((string)$cols[$headerMap['jp']]) : '';
            $en = isset($cols[$headerMap['en']]) ? trim((string)$cols[$headerMap['en']]) : '';

            if ($jp === '' && $en === '') continue;

            $rows[] = [
                'jp' => $jp,
                'en' => $en,
            ];
        }

        fclose($handle);

        if (empty($rows)) return [[], 'CSVからデータを取得できませんでした。'];

        update_option(self::OPT_CACHE, [
            'mtime' => (int)$mtime,
            'rows'  => $rows,
        ], false);

        return [$rows, ''];
    }

    private static function pick_today_row($rows) {
        $count = count($rows);
        if ($count === 0) return null;

        // 日付をseedにした“日替わりランダム”（同日同結果）
        $seed = crc32(wp_date('Y-m-d'));
        $x = ($seed * 1103515245 + 12345) & 0x7fffffff;
        $index = $x % $count;

        return $rows[$index];
    }

    public static function shortcode($atts) {
        [$rows, $err] = self::load_rows();

        if ($err) {
            if (current_user_can('manage_options')) {
                return '<div style="padding:12px;border:1px solid #f0c2c2;border-radius:10px;">'
                    . '<strong>Daily Quotes CSV:</strong> ' . esc_html($err)
                    . '</div>';
            }
            return '';
        }

        $row = self::pick_today_row($rows);
        if (!$row) return '';

        $jp = $row['jp'] ?? '';
        $en = $row['en'] ?? '';

        $html  = '<div class="daily-positive-word" style="padding:14px 16px;border:1px solid #ddd;border-radius:10px;line-height:1.8;">';
        $html .= '<div style="font-size:12px;margin-bottom:6px;">
		<span style="font-weight:600; opacity:.6;">- A Word for Today -</span>
            <span style="opacity:.6;">' . esc_html(wp_date('Y-m-d')) . '</span>
          </div>';

        if ($jp !== '') {
            $html .= '<div style="font-size:18px;font-weight:600;">' . esc_html($jp) . '</div>';
        }
        if ($en !== '') {
            $html .= '<div style="font-size:14px;opacity:.75;margin-top:6px;">' . esc_html($en) . '</div>';
        }

        $html .= '</div>';
        return $html;
    }

    public static function inject_front_page_quote($content) {
        if (!is_front_page() || is_admin()) return $content;
        if (has_shortcode($content, 'daily_positive_word')) return $content;
        return do_shortcode('[daily_positive_word]') . "\n\n" . $content;
    }
}

Daily_Positive_Word::init();

