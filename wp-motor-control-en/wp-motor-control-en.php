<?php
/**
 * Plugin Name:  Motor Control EN – The Silent Key Factor in Sports Performance
 * Description:  English branching scenario based on the article by Raquel Font-Lladó (UdG). Add [motor_control_en] to any page.
 * Version:      1.0.0
 * Requires PHP: 7.4
 * License:      GPL-2.0-or-later
 * Text Domain:  motor-control-en
 */

if ( ! defined( 'ABSPATH' ) ) exit;

define( 'MC_EN_VERSION',    '1.0.0' );
define( 'MC_EN_PLUGIN_DIR', plugin_dir_path( __FILE__ ) );
define( 'MC_EN_PLUGIN_URL', plugin_dir_url( __FILE__ ) );

class MotorControlEnPlugin {

    public function __construct() {
        add_shortcode( 'motor_control_en',  [ $this, 'render_shortcode' ] );
        add_action( 'rest_api_init',        [ $this, 'register_rest_routes' ] );
        add_action( 'admin_menu',           [ $this, 'add_admin_menu' ] );
    }

    /* ── Shortcode ─────────────────────────────────────── */

    public function render_shortcode( $atts ) {
        wp_enqueue_style(
            'mc-en-style',
            MC_EN_PLUGIN_URL . 'assets/css/style.css',
            [],
            MC_EN_VERSION
        );
        wp_enqueue_script(
            'mc-en-progress',
            MC_EN_PLUGIN_URL . 'assets/js/progress.js',
            [],
            MC_EN_VERSION,
            true
        );
        wp_localize_script( 'mc-en-progress', 'MC_EN_Config', [
            'restUrl'    => esc_url_raw( rest_url( 'motor-control-en/v1/' ) ),
            'nonce'      => wp_create_nonce( 'wp_rest' ),
            'isLoggedIn' => is_user_logged_in(),
        ] );
        wp_enqueue_script(
            'mc-en-scenario',
            MC_EN_PLUGIN_URL . 'assets/js/scenario.js',
            [ 'mc-en-progress' ],
            MC_EN_VERSION,
            true
        );
        wp_add_inline_script(
            'mc-en-scenario',
            "(function(){ if(document.readyState!=='loading'){ Engine.init(); } else { document.addEventListener('DOMContentLoaded',function(){ Engine.init(); }); } })();"
        );

        ob_start();
        include MC_EN_PLUGIN_DIR . 'templates/course-template.php';
        return ob_get_clean();
    }

    /* ── REST API ───────────────────────────────────────── */

    public function register_rest_routes() {
        register_rest_route( 'motor-control-en/v1', '/progress', [
            [
                'methods'             => WP_REST_Server::READABLE,
                'callback'            => [ $this, 'get_progress' ],
                'permission_callback' => [ $this, 'require_login' ],
            ],
            [
                'methods'             => WP_REST_Server::CREATABLE,
                'callback'            => [ $this, 'save_progress' ],
                'permission_callback' => [ $this, 'require_login' ],
                'args' => [
                    'progress_data' => [
                        'required'          => true,
                        'type'              => 'string',
                        'sanitize_callback' => 'sanitize_text_field',
                    ],
                ],
            ],
        ] );
    }

    public function require_login() {
        return is_user_logged_in();
    }

    public function get_progress( WP_REST_Request $request ) {
        $data = get_user_meta( get_current_user_id(), 'mc_en_progress', true );
        return rest_ensure_response( [ 'data' => $data ?: '' ] );
    }

    public function save_progress( WP_REST_Request $request ) {
        update_user_meta(
            get_current_user_id(),
            'mc_en_progress',
            $request->get_param( 'progress_data' )
        );
        return rest_ensure_response( [ 'success' => true ] );
    }

    /* ── Admin page ─────────────────────────────────────── */

    public function add_admin_menu() {
        add_options_page(
            'Motor Control EN',
            'Motor Control EN',
            'manage_options',
            'motor-control-en',
            [ $this, 'admin_page' ]
        );
    }

    public function admin_page() {
        global $wpdb;
        $count = (int) $wpdb->get_var(
            "SELECT COUNT(*) FROM {$wpdb->usermeta} WHERE meta_key = 'mc_en_progress'"
        );
        ?>
        <div class="wrap">
            <h1>Motor Control EN – The Silent Key Factor in Sports Performance</h1>
            <p>Based on the article by Raquel Font-Lladó (University of Girona). English version with Viktor as researcher.</p>
            <p>Add <code>[motor_control_en]</code> to any page to display the course.</p>
            <table class="form-table" role="presentation">
                <tr>
                    <th scope="row">Students with saved progress</th>
                    <td><strong><?php echo esc_html( $count ); ?></strong></td>
                </tr>
                <tr>
                    <th scope="row">Plugin version</th>
                    <td><?php echo esc_html( MC_EN_VERSION ); ?></td>
                </tr>
            </table>
        </div>
        <?php
    }
}

new MotorControlEnPlugin();
