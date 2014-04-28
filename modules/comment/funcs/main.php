<?php

/**
 * @Project NUKEVIET 4.x
 * @Author VINADES.,JSC (contact@vinades.vn)
 * @Copyright (C) 2014 VINADES.,JSC. All rights reserved
 * @License GNU/GPL version 2 or any later version
 * @Createdate Mon, 27 Jan 2014 00:08:04 GMT
 */

if( ! defined( 'NV_IS_MOD_COMMENT' ) ) die( 'Stop!!!' );

$module = $nv_Request->get_string( 'module', 'post,get' );

if( ! empty( $module ) AND isset( $module_config[$module]['activecomm'] ) )
{
	// Kiểm tra module có được Sử dụng chức năng bình luận

	$area = $nv_Request->get_int( 'area', 'post,get', 0 );
	$id = $nv_Request->get_int( 'id', 'post,get', 0 );
	$view_comm = $nv_Request->get_int( 'view', 'post,get', 0 );
	$allowed_comm = $nv_Request->get_int( 'allowed', 'post,get', 0 );
	$checkss = $nv_Request->get_string( 'checkss', 'post,get' );
	$page = $nv_Request->get_int( 'page', 'get', 0 );

	if( $id > 0 AND $module_config[$module]['activecomm'] == 1 AND $checkss == md5( $module . '-' . $area . '-' . $id . '-' . $view_comm . '-' . $allowed_comm . '-' . NV_CACHE_PREFIX ) )
	{
		$base_url = NV_BASE_SITEURL . 'index.php?' . NV_LANG_VARIABLE . '=' . NV_LANG_DATA . '&amp;' . NV_NAME_VARIABLE . '=comment&module=' . $module . '&area=' . $area . '&id=' . $id . '&view=' . $view_comm . '&allowed=' . $allowed_comm . '&checkss=' . $checkss;

		// Kiểm tra quyền xem bình luận
		$form_login = 0;
		$view = intval( $module_config[$module]['view_comm'] );
		if( $view == 3 )
		{
			// Quyền hạn xem bình luận theo bài viết
			$view = $view_comm;
		}
		if( $view == 1 or ( $view == 2 and defined( 'NV_IS_USER' ) ) )
		{
			$view_comm = 1;
		}
		else
		{
			$view_comm = 0;
		}

		// Kiểm tra quyền đăng bình luận
		$allowed = intval( $module_config[$module]['allowed_comm'] );
		if( $allowed == 3 )
		{
			// Quyền hạn đăng bình luận theo bài viết
			$allowed = $allowed_comm;
		}
		if( $allowed == 1 or ( $allowed == 2 and defined( 'NV_IS_USER' ) ) )
		{
			$allowed_comm = 1;
		}
		else
		{
			$allowed_comm = 0;
		}

		if( ( $view == 2 OR $allowed == 2 ) AND ! defined( 'NV_IS_USER' ) )
		{
			$form_login = 1;
		}

		$array_data = array();

		$page_title = $module_info['custom_title'];
		$key_words = $module_info['keywords'];
		$global_config['mudim_active'] = 0;

		$sortcomm_old = $nv_Request->get_int( 'sortcomm', 'cookie', $module_config[$module]['sortcomm'] );
		$sortcomm = $nv_Request->get_int( 'sortcomm', 'post,get', $sortcomm_old );
		if( $sortcomm < 0 OR $sortcomm > 2 )
		{
			$sortcomm = 0;
		}
		if( $sortcomm_old != $sortcomm )
		{
			$nv_Request->set_Cookie( 'sortcomm', $sortcomm, NV_LIVE_COOKIE_TIME );
		}

		$is_delete = false;
		if( defined( 'NV_IS_ADMIN' ) )
		{
			$is_delete = true;
		}
		elseif( defined( 'NV_IS_ADMIN' ) )
		{
			$adminscomm = explode( ',', $module_config[$module]['adminscomm'] );
			if( in_array( $admin_info['admin_id'], $adminscomm ) )
			{
				$is_delete = true;
			}
		}

		if( $view_comm )
		{
			$comment_array = nv_comment_data( $module, $area, $id, $allowed_comm, $page, $sortcomm, $base_url );
			$comment = nv_comment_theme( $module, $comment_array, $is_delete );
		}
		else
		{
			$comment = '';
		}

		$contents = nv_theme_comment_main( $module, $area, $id, $view_comm, $allowed_comm, $checkss, $comment, $sortcomm, $base_url, $form_login );
		include NV_ROOTDIR . '/includes/header.php';
		echo nv_site_theme( $contents, false );
		include NV_ROOTDIR . '/includes/footer.php';
	}
	else
	{
		die( 'Stop!!!' );
	}
}

Header( 'Location: ' . NV_BASE_SITEURL . 'index.php?' . NV_LANG_VARIABLE . '=' . NV_LANG_DATA );
die();