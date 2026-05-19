<?php
/**
 * Plugin Name:  Curs Interactiu – Entrenador Profesional
 * Description:  Escenari ramificat gamificat (CD Olimpia). Afegeix [curs_interactiu] a qualsevol pàgina.
 * Version:      1.0.0
 * Requires PHP: 7.4
 * License:      GPL-2.0-or-later
 * Text Domain:  curs-interactiu
 */

if ( ! defined( 'ABSPATH' ) ) exit;

define( 'CI_VERSION',    '1.0.0' );
define( 'CI_PLUGIN_DIR', plugin_dir_path( __FILE__ ) );
define( 'CI_PLUGIN_URL', plugin_dir_url( __FILE__ ) );

class CursInteractiu {

    public function __construct() {
        add_shortcode( 'curs_interactiu', [ $this, 'render_shortcode' ] );
        add_action( 'rest_api_init',      [ $this, 'register_rest_routes' ] );
        add_action( 'admin_menu',         [ $this, 'add_admin_menu' ] );
    }

    /* ── Shortcode ─────────────────────────────────────── */

    public function render_shortcode( $atts ) {
        wp_enqueue_style(
            'ci-style',
            CI_PLUGIN_URL . 'assets/css/style.css',
            [],
            CI_VERSION
        );
        wp_enqueue_script(
            'ci-progress',
            CI_PLUGIN_URL . 'assets/js/progress.js',
            [],
            CI_VERSION,
            true
        );
        wp_localize_script( 'ci-progress', 'CI_Config', [
            'restUrl'    => esc_url_raw( rest_url( 'curs-interactiu/v1/' ) ),
            'nonce'      => wp_create_nonce( 'wp_rest' ),
            'isLoggedIn' => is_user_logged_in(),
        ] );
        wp_enqueue_script(
            'ci-scenario',
            CI_PLUGIN_URL . 'assets/js/scenario.js',
            [ 'ci-progress' ],
            CI_VERSION,
            true
        );
        /* Init Engine after both scripts are loaded in footer */
        wp_add_inline_script(
            'ci-scenario',
            "(function(){ if(document.readyState!=='loading'){ Engine.init(); } else { document.addEventListener('DOMContentLoaded',function(){ Engine.init(); }); } })();"
        );

        ob_start();
        include CI_PLUGIN_DIR . 'templates/course-template.php';
        return ob_get_clean();
    }

    /* ── REST API ───────────────────────────────────────── */

    public function register_rest_routes() {
        register_rest_route( 'curs-interactiu/v1', '/progress', [
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
        $data = get_user_meta( get_current_user_id(), 'ci_entrenador_progress', true );
        return rest_ensure_response( [ 'data' => $data ?: '' ] );
    }

    public function save_progress( WP_REST_Request $request ) {
        update_user_meta(
            get_current_user_id(),
            'ci_entrenador_progress',
            $request->get_param( 'progress_data' )
        );
        return rest_ensure_response( [ 'success' => true ] );
    }

    /* ── Admin page ─────────────────────────────────────── */

    public function add_admin_menu() {
        add_options_page(
            'Curs Interactiu',
            'Curs Interactiu',
            'manage_options',
            'curs-interactiu',
            [ $this, 'admin_page' ]
        );
    }

    public function admin_page() {
        global $wpdb;
        $count = (int) $wpdb->get_var(
            "SELECT COUNT(*) FROM {$wpdb->usermeta} WHERE meta_key = 'ci_entrenador_progress'"
        );
        ?>
        <div class="wrap">
            <h1>Curs Interactiu – Dimensiones del Entrenador Profesional</h1>
            <p>Afegeix <code>[curs_interactiu]</code> a qualsevol pàgina per mostrar el curs.</p>
            <table class="form-table" role="presentation">
                <tr>
                    <th scope="row">Alumnes amb progrés guardat</th>
                    <td><strong><?php echo esc_html( $count ); ?></strong></td>
                </tr>
                <tr>
                    <th scope="row">Versió del plugin</th>
                    <td><?php echo esc_html( CI_VERSION ); ?></td>
                </tr>
            </table>
        </div>
        <?php
    }
}

new CursInteractiu();
